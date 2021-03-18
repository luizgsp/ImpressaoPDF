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

namespace GeracaoColetasPDF
{
    public class Program
    {
        //private StreamReader streamToPrint;
        private static string _primeiraLetra;
        private static string PrimeiraLetra
        {
            get { if (_primeiraLetra == null) { return ""; } return _primeiraLetra; }
            set { _primeiraLetra = value; }
        }
        private static string _impressora1;
        private static string Impressora1
        {
            get { if (_impressora1 == null) { return ""; } return _impressora1; }
            set { _impressora1 = value; }
        }
        private static string _impressora2;
        private static string Impressora2
        {
            get { if (_impressora2 == null) { return ""; } return _impressora2; }
            set { _impressora2 = value; }
        }

        private static bool _imprime;
        private static bool Imprime
        {
            get { return _imprime; }
            set { _imprime = value; }
        }

        private static string _caminhoOrigem;
        private static string CaminhoOrigem
        {
            get { if (_caminhoOrigem == null) { return ""; } return _caminhoOrigem; }
            set { _caminhoOrigem = value; }
        }
        private static string _caminhoDestino;
        private static string CaminhoDestino
        {
            get { if (_caminhoDestino == null) { return ""; } return _caminhoDestino; }
            set { _caminhoDestino = value; }
        }

        private static string _servidorSMTP;
        private static string ServidorSMTP
        {
            get { if (_servidorSMTP == null) { return ""; } return _servidorSMTP; }
            set { _servidorSMTP = value; }
        }

        private static string _contaEmail;
        private static string ContaEmail
        {
            get { if (_contaEmail == null) { return ""; } return _contaEmail; }
            set { _contaEmail = value; }
        }

        private static string _senha;
        private static string Senha
        {
            get { if (_senha == null) { return ""; } return _senha; }
            set { _senha = value; }
        }

        private static string _portaSMTP;
        private static string PortaSMTP
        {
            get { if (_portaSMTP == null) { return ""; } return _portaSMTP; }
            set { _portaSMTP = value; }
        }

        private static string _vendedor;
        private static string Vendedor
        {
            get { if (_vendedor == null) { return ""; } return _vendedor; }
            set { _vendedor = value; }
        }

        private static string _comprador;
        private static string Comprador
        {
            get { if (_comprador == null) { return ""; } return _comprador; }
            set { _comprador = value; }
        }

        private static string _cliente;
        private static string Cliente
        {
            get { if (_cliente == null) { return ""; } return _cliente; }
            set { _cliente = value; }
        }

        private static string _cotacao;
        private static string Cotacao
        {
            get { if (_cotacao == null) { return ""; } return _cotacao; }
            set { _cotacao = value; }
        }

        private static string _cnpjCpf;
        private static string CnpjCpf
        {
            get { if (_cnpjCpf == null) { return ""; } return _cnpjCpf; }
            set { _cnpjCpf = value; }
        }

        private static string _emailCli;
        private static string EmailCli
        {
            get { if (_emailCli == null) { return ""; } return _emailCli; }
            set { _emailCli = value; }
        }

        private static string _corpoEmail;
        private static string CorpoEmail
        {
            get { if (_corpoEmail == null) { return ""; } return _corpoEmail; }
            set { _corpoEmail = value; }
        }

        private static string _observacao1;
        private static string Observacao1
        {
            get { if (_observacao1 == null) { return ""; } return _observacao1; }
            set { _observacao1 = value; }
        }

        private static string _observacao2;
        private static string Observacao2
        {
            get { if (_observacao2 == null) { return ""; } return _observacao2; }
            set { _observacao2 = value; }
        }

        private static BaseColor _CorFonte;
        private static BaseColor CorFonte { get { return _CorFonte; } set { _CorFonte = value; } }

        private static bool _AlteraCor;
        private static bool AlteraCor { get { return _AlteraCor; } set { _AlteraCor = value; } }
        private static int Count = 0;


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
                //Console.WriteLine(@"Relatorio impresso {0}", RelatorioPDF);
            }
        }

        private static int ListaArquivos()
        {

            System.Threading.Thread.Sleep(2000);
            string CaminhoXML = Directory.GetCurrentDirectory() + @"\Config.xml";
            DataSet Ds = new DataSet();
            Ds.ReadXml(CaminhoXML);
            DataTable Dt = Ds.Tables[0];
            DataRow Dr = Dt.Rows[0];
            if (Dr != null)
            {
                CaminhoOrigem = Dr["CaminhoOrigem"].ToString();
                CaminhoDestino = Dr["CaminhoDestino"].ToString();
                ServidorSMTP = Dr["ServidorSMTP"].ToString();
                //ContaEmail = Dr["ContaEmail"].ToString();
                //Senha = Dr["Pass"].ToString();
                PortaSMTP = Dr["PortaSMTP"].ToString();
                Impressora1 = Dr["CaminhoImpressora1"].ToString();
                Impressora2 = Dr["CaminhoImpressora2"].ToString();
                CorpoEmail = Dr["CorpoEmail"].ToString();
                Console.Clear();
                Console.WriteLine("\t\t\tMinas Ferramentas Ltda.");
                Console.WriteLine();
                Console.WriteLine("\t      Impressao e Envio de emails de Coletas em PDF");
                Console.WriteLine();
                Console.WriteLine("\t     Belo Horizonte, {0}", DateTime.Now.ToString("dd 'de' MMMM 'de' yyyy HH:mm:ss"));
                Console.WriteLine();
                //Console.WriteLine(@"Caminho de Origem..: {0}", CaminhoOrigem);
                //Console.WriteLine();
                //Console.WriteLine(@"Caminho de destino.: {0}", CaminhoDestino);

            }

            DirectoryInfo diretorio = new DirectoryInfo(CaminhoOrigem);
            if (!diretorio.Exists)
            {
                Console.WriteLine("Não encontrado o diretorio especificado!");
                Imprime = false;
                return 0;
            }

            //System.Threading.Thread.Sleep(1500);
            FileInfo[] Arquivos = diretorio.GetFiles("*.txt");
            Console.WriteLine();
            Console.WriteLine(@"Arquivos.: {0}", Arquivos.Count());
            Console.WriteLine();
            if (Arquivos.Count() == 0)
            {
                try
                {
                    Arquivos = diretorio.GetFiles("*.pdf");
                    if (Arquivos.Count() > 0)
                    {
                        foreach (FileInfo fileinfo in Arquivos)
                        {
                            System.Threading.Thread.Sleep(5000);
                            System.IO.File.Move(CaminhoOrigem + fileinfo, CaminhoDestino + fileinfo);
                            Console.WriteLine(@"Movido o Arquivo {0} para a pasta Impressos.", fileinfo);

                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    System.Threading.Thread.Sleep(10000);

                }
            }
            else
            {
                foreach (FileInfo fileinfo in Arquivos)
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
            }
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("\t\t\t\tAguarde....");
            // if (Imprime) { Close(); }
            return 1;
        }

        private static void GerarPDF(string ColetaTXT)
        {
            Document documento = new Document(PageSize.A4, 5, 5, 20, 20);
            documento.AddAuthor("Minas Ferramentas Ltda");
            documento.AddSubject("Cotação " + ColetaTXT.Replace(".txt", "").Replace(".TXT", ""));
            documento.SetPageSize(PageSize.A4.Rotate());
            int CodVendedor = 0;  // Coleta.Replace(".txt", "").Replace(".TXT", "");
            //Vendedor = Vendedor.Substring(5, 3);
            DateTime Dt1 = DateTime.Now;
            string ColetaPDF = "MFL Cotação nº " + ColetaTXT.Replace(".txt", "").Replace(".TXT", "");
            ColetaPDF = ColetaPDF + " em " + Dt1.ToString("dd-MM-yyyy HH-mm-ss-fff") + ".pdf";
            string ColetaPDFAenviar = ColetaPDF;
            if (System.IO.File.Exists(CaminhoOrigem + ColetaPDF))
            {
                System.IO.File.Delete(ColetaPDF);
            }
            ColetaPDF = CaminhoOrigem + ColetaPDF;
            try
            {

                PdfWriter Writer = PdfWriter.GetInstance(documento, new FileStream(ColetaPDF, FileMode.Create));

                //Seleciona o arquivo para a imagem de marca d'água
                string MarcaDagua = Directory.GetCurrentDirectory() + @"\LOGO_MF_SELO_PB.jpg";
                iTextSharp.text.Image ImgMarcaDagua = iTextSharp.text.Image.GetInstance(MarcaDagua);
                //Informa a posição da Marca d'Água
                ImgMarcaDagua.SetAbsolutePosition(150, 500);

                //Abre o documento
                documento.Open();
                documento.NewPage();
                iTextSharp.text.Rectangle page = documento.PageSize;

                //Adiciona a imagem ao documento
                //documento.Add(ImgMarcaDagua);

                PdfPTable table = new PdfPTable(1);
                table.WidthPercentage = 100;

                //IMPRIMINDO O CABECALHO MINAS FERRAMENTAS LTDA
                PdfPTable TableCabec = new PdfPTable(6);
                TableCabec.WidthPercentage = 98;

                string LogoMarca = Directory.GetCurrentDirectory() + @"\LogoMF2.jpg";
                iTextSharp.text.Image ImgMinas = iTextSharp.text.Image.GetInstance(LogoMarca);
                ImgMinas.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                //ImgMinas.Top = iTextSharp.text.Image.ALIGN_CENTER;

                string LogoFluke = Directory.GetCurrentDirectory() + @"\IMGFluke.jpg";
                iTextSharp.text.Image ImgFluke = iTextSharp.text.Image.GetInstance(LogoFluke);
                ImgFluke.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                ImgFluke.ScaleAbsolute(100f, 50f);
                //ImgFluke.Top = iTextSharp.text.Image.ALIGN_CENTER;

                PdfPTable tblCabec = new PdfPTable(13);
                tblCabec.WidthPercentage = 100;
                PdfPTable tableClientes = new PdfPTable(1);
                PdfPTable tableClientes2 = new PdfPTable(6);
                PdfPTable tableItens = new PdfPTable(15);
                tableItens.WidthPercentage = 95;

                PdfPTable tableTotais = new PdfPTable(5);
                PdfPTable tableObs = new PdfPTable(6);
                PdfPTable tableObs2 = new PdfPTable(2);


                PdfPCell cell;
                PdfPCell cellCabec = new PdfPCell(ImgMinas);
                cellCabec.Rowspan = 4;
                cellCabec.Colspan = 2;
                cellCabec.Border = 0;
                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                cellCabec.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                TableCabec.AddCell(cellCabec);

                cellCabec = new PdfPCell(ImgFluke);
                cellCabec.Rowspan = 4;
                cellCabec.Colspan = 2;
                cellCabec.Border = 0;
                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                cellCabec.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                TableCabec.AddCell(cellCabec);

                string[] arrValores = new string[2];
                string LetraAnt = "";
                string Detalhes = "";
                string[] linhaA1 = new string[10];
                int ContaLinhas = 0;
                int intLinha = 0;
                int ind1 = 6;
                bool Linha1 = true;
                //Boolean PulaPagina = false;
                string PrimeiroItemCabec = "";
                AlteraCor = true;
                CorFonte = BaseColor.BLACK;
                Vendedor = "";
                Comprador = "";
                Cliente = "";
                Cotacao = "";
                CnpjCpf = "";
                EmailCli = "";
                Observacao1 = "";
                Observacao2 = "";
                int TotalPaginas = ContadorPaginas(CaminhoOrigem + ColetaTXT);
                int ContaPaginas = 1;
                StreamReader srArquivo = new StreamReader(CaminhoOrigem + ColetaTXT);
                while (srArquivo.Peek() != -1)
                {
                    string strLinha = srArquivo.ReadLine();
                    if (strLinha != "")
                    {
                        int TamanhoLinha = strLinha.Length;
                        string Letra = strLinha.Substring(0, 1);
                        if ((TamanhoLinha == 4) && (Linha1))
                        {
                            PrimeiraLetra = strLinha.Substring(0, 1);
                            CodVendedor = int.Parse(strLinha.Substring(1, 3));
                            Linha1 = false;
                        }
                        else
                        {
                            if (TamanhoLinha > 1)
                            {
                                Detalhes = strLinha.Substring(1, (TamanhoLinha - 1)); //Conteúdo
                                switch (Letra)
                                {
                                    case "A"://imprime o cabecalho fixo
                                        string[] linhaA = Detalhes.Split(char.Parse(":"));

                                        switch (intLinha)
                                        {
                                            case 0:// Numero da Cotacao
                                                //cellCabec = new PdfPCell(new Phrase(@"vendas@minasferramentas.com.br", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL)));
                                                //cellCabec.Border = 0;
                                                //cellCabec.Colspan = 2;
                                                //cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                //cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                //cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                //TableCabec.AddCell(cellCabec);

                                                Chunk c5 = new Chunk(linhaA[0], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 16, iTextSharp.text.Font.NORMAL));
                                                Chunk c6 = new Chunk(linhaA[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 16, iTextSharp.text.Font.NORMAL));
                                                Phrase p2 = new Phrase();
                                                p2.Add(c5);
                                                p2.Add(c6);
                                                cellCabec = new PdfPCell(p2);
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);
                                                linhaA1[0] = Detalhes;
                                                Cotacao = Detalhes;
                                                break;
                                            case 1://Data da cotacao e pagina
                                                //cellCabec = new PdfPCell(new Phrase(@"CNPJ.: 17.194.994/0001-27 IE.: 0620080420094", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL)));
                                                //cellCabec.Border = 0;
                                                //cellCabec.Colspan = 2;
                                                //cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                //cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                //cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                //TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(Detalhes + @" Pag.: " + ContaPaginas.ToString("##0") + @"/" + TotalPaginas, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);
                                                linhaA1[1] = Detalhes;
                                                break;
                                            case 2://Vendedor
                                                //cellCabec = new PdfPCell(new Phrase(@"Av. Bias Fortes, 1853 - Barro Preto - Belo Horizonte/MG",
                                                //                FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                //cellCabec.Border = 0;
                                                //cellCabec.Colspan = 2;
                                                //cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                //cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                //cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                //TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(Detalhes, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);
                                                Vendedor = linhaA[1];
                                                linhaA1[2] = Detalhes;
                                                break;
                                            case 3://Referente
                                                //cellCabec = new PdfPCell(new Phrase(@"Cep:30.170-012   Fone: (31) 2101-6000 Fax: (31) 2101-6010", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                //cellCabec.Border = 0;
                                                //cellCabec.Colspan = 2;
                                                //cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                //cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                //cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                //TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(Detalhes, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);
                                                linhaA1[3] = Detalhes;
                                                break;
                                            case 4:
                                                cellCabec = new PdfPCell(new Phrase(Detalhes, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.BOLDITALIC)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 6;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);
                                                linhaA1[4] = Detalhes;
                                                break;

                                            default:
                                                cellCabec = new PdfPCell(new Phrase(@"Belo Horizonte-MG CEP 30.170-012", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(@"Tel.: (31) 2101.6000 - Fax: (31) 2101.6010", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);

                                                table.AddCell(TableCabec);
                                                linhaA1[4] = Detalhes;
                                                break;
                                        }
                                        intLinha++;
                                        break;
                                    case "B":
                                        if (LetraAnt != Letra) { LetraAnt = Letra; }
                                        if (Detalhes.Length > 2)
                                        {
                                            string[] linhaB = Detalhes.Split(char.Parse("|"));
                                            switch (linhaB.Length)
                                            {
                                                case 3: // Razão Social, CNPJ, IE
                                                    foreach (string dados in linhaB)
                                                    {
                                                        string[] linhaB1 = dados.Split(char.Parse(":"));
                                                        if (linhaB1.Length == 2)
                                                        {

                                                            Chunk c1 = new Chunk(linhaB1[0], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL));
                                                            Chunk c2 = new Chunk(linhaB1[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, iTextSharp.text.Font.NORMAL));
                                                            Phrase NovaLinha = new Phrase(Environment.NewLine);
                                                            Paragraph paragrafo = new Paragraph();
                                                            paragrafo.Add(c1);
                                                            paragrafo.Add(NovaLinha);
                                                            paragrafo.Add(c2);
                                                            cellCabec = new PdfPCell(paragrafo);
                                                            //cellCabec.Border = 0;
                                                            if (linhaB1[0].Contains("Razao Social"))
                                                            {
                                                                cellCabec.Colspan = 4;
                                                                Cliente = linhaB1[1];
                                                            }
                                                            if (linhaB1[0].Contains("CNPJ"))
                                                                CnpjCpf = linhaB1[1];
                                                            cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                            cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                            cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                            tableClientes2.AddCell(cellCabec);
                                                        }
                                                    }
                                                    linhaA1[5] = Detalhes;
                                                    break;
                                                case 6: // Logradouro, Cidade, UF, Cep, Tel, Fax
                                                    foreach (string dados in linhaB)
                                                    {
                                                        string[] linhaB2 = dados.Split(char.Parse(":"));
                                                        if (linhaB2.Length == 2)
                                                        {
                                                            Phrase c1 = new Phrase(linhaB2[0], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL));
                                                            Phrase c2 = new Phrase(linhaB2[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, iTextSharp.text.Font.NORMAL));
                                                            Phrase NovaLinha = new Phrase(Environment.NewLine);

                                                            Paragraph paragrafo = new Paragraph();
                                                            paragrafo.Add(c1);
                                                            paragrafo.Add(NovaLinha);
                                                            paragrafo.Add(c2);
                                                            cellCabec = new PdfPCell(paragrafo);
                                                            //cellCabec.Border = 0;
                                                            if (linhaB2[0].Contains("Logradouro"))
                                                                cellCabec.Colspan = 3;
                                                            cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                            cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                            cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                            tableClientes2.AddCell(cellCabec);
                                                        }
                                                    }
                                                    linhaA1[6] = Detalhes;
                                                    ind1 = 6;
                                                    break;
                                                case 1: // Email, Att e Mensagem
                                                    string[] linhaB3 = Detalhes.Split(char.Parse(":"));
                                                    if (linhaB3.Length == 2)
                                                    {
                                                        Phrase c1 = new Phrase(linhaB3[0], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL));
                                                        Phrase c2 = new Phrase(linhaB3[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, iTextSharp.text.Font.NORMAL));
                                                        Phrase NovaLinha = new Phrase(Environment.NewLine);
                                                        Paragraph paragrafo = new Paragraph();
                                                        paragrafo.Add(c1);
                                                        paragrafo.Add(NovaLinha);
                                                        paragrafo.Add(c2);
                                                        cellCabec = new PdfPCell(paragrafo);

                                                        if (linhaB3[0].Contains("E-mail"))
                                                            EmailCli = linhaB3[1];
                                                        if (linhaB3[0].Contains("ATT"))
                                                            Comprador = linhaB3[1];

                                                        //cellCabec.Border = 0;
                                                        cellCabec.Colspan = 2;
                                                        cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                        cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                        cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                        tableClientes2.AddCell(cellCabec);
                                                        ind1++;
                                                        linhaA1[ind1] = Detalhes;
                                                    }
                                                    else
                                                    {
                                                        cellCabec = new PdfPCell(new Phrase(Detalhes, FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL)));
                                                        cellCabec.Border = 0;
                                                        cellCabec.Colspan = 6;
                                                        cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                        cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                        cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                        tableClientes2.AddCell(cellCabec);
                                                        ind1++;
                                                        linhaA1[ind1] = Detalhes;
                                                    }
                                                    break;

                                                default:
                                                    cellCabec = new PdfPCell(new Phrase(Detalhes, FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL)));
                                                    cellCabec.Border = 0;
                                                    cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                    cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                    cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableClientes2.AddCell(cellCabec);
                                                    ind1++;
                                                    linhaA1[ind1] = Detalhes;
                                                    break;
                                            }
                                        }
                                        break;
                                    case "C":
                                        if (LetraAnt != Letra)  // IMPRESSAO DO CABECALHO DOS ITENS
                                        {
                                            //TableCabec
                                            cell = new PdfPCell(TableCabec);
                                            cell.Colspan = 15;
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            tableItens.AddCell(cell);

                                            //tableClientes2
                                            cell = new PdfPCell(tableClientes2);
                                            cell.Colspan = 15;
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            tableItens.AddCell(cell);

                                            //tableItens
                                            table.AddCell(tableItens);


                                            LetraAnt = Letra;
                                            ContaLinhas = 0;
                                        }

                                        /* COLOCAR CABEÇALHO PARA CADA PAGINA E CONTADOR DE PAGINAS*/

                                        ContaLinhas++;
                                        if (ContaLinhas > 28) /*ATENCAO: Quando trocar a QUANTIDADE de linhas mudar tambem no prog. MF665B na linha 1818*/
                                        {
                                            AlteraCor = true;
                                            ContaLinhas = 1;
                                            //PulaPagina = true;
                                            ContaPaginas++;
                                            TableCabec = new PdfPTable(6);
                                            TableCabec.WidthPercentage = 100;
                                            tableClientes2 = new PdfPTable(6);
                                            tableClientes2.WidthPercentage = 100;

                                            LogoMarca = Directory.GetCurrentDirectory() + @"\LogoMF2.jpg";
                                            ImgMinas = iTextSharp.text.Image.GetInstance(LogoMarca);
                                            ImgMinas.Alignment = iTextSharp.text.Image.ALIGN_CENTER;


                                            cellCabec = new PdfPCell(ImgMinas);
                                            cellCabec.Rowspan = 4;
                                            cellCabec.Colspan = 2;
                                            cellCabec.Border = 0;
                                            cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                            cellCabec.HorizontalAlignment = PdfPCell.ALIGN_MIDDLE;
                                            TableCabec.AddCell(cellCabec);

                                            cellCabec = new PdfPCell(new Phrase(@"vendas@minasferramentas.com.br", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                            TableCabec.AddCell(cellCabec);
                                            string[] linhaC = linhaA1[0].Split(char.Parse(":"));

                                            Chunk c5 = new Chunk(linhaC[0], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 16, iTextSharp.text.Font.NORMAL));
                                            Chunk c6 = new Chunk(linhaC[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 16, iTextSharp.text.Font.NORMAL));
                                            Phrase p2 = new Phrase();
                                            p2.Add(c5);
                                            p2.Add(c6);
                                            cellCabec = new PdfPCell(p2);
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                            TableCabec.AddCell(cellCabec);


                                            cellCabec = new PdfPCell(new Phrase(@"CNPJ.: 17.194.994/0001-27 IE.: 0620080420094", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                            TableCabec.AddCell(cellCabec);

                                            cellCabec = new PdfPCell(new Phrase(linhaA1[1] + @" Pag.: " + ContaPaginas.ToString("##0") + @"/" + TotalPaginas, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                            TableCabec.AddCell(cellCabec);

                                            cellCabec = new PdfPCell(new Phrase(@"Av. Bias Fortes, 1853 - Barro Preto - Belo Horizonte/MG",
                                                            FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                            TableCabec.AddCell(cellCabec);

                                            cellCabec = new PdfPCell(new Phrase(linhaA1[2], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                            TableCabec.AddCell(cellCabec);

                                            cellCabec = new PdfPCell(new Phrase(@"Cep:30.170-012   Fone: (31) 2101-6000 Fax: (31) 2101-6010", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                            TableCabec.AddCell(cellCabec);

                                            cellCabec = new PdfPCell(new Phrase(linhaA1[3], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                            TableCabec.AddCell(cellCabec);
                                            /*----------------------------------------------------------------------------------------------------------------------------------*/
                                            string[] linhaC1 = linhaA1[5].Split(char.Parse("|")); // Razão Social, CNPJ, IE
                                            foreach (string dados in linhaC1)
                                            {
                                                string[] linhaB1 = dados.Split(char.Parse(":"));
                                                if (linhaB1.Length == 2)
                                                {
                                                    Chunk c1 = new Chunk(linhaB1[0], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL));
                                                    Chunk c2 = new Chunk(linhaB1[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, iTextSharp.text.Font.NORMAL));
                                                    Phrase NovaLinha = new Phrase(Environment.NewLine);
                                                    Paragraph paragrafo = new Paragraph();
                                                    paragrafo.Add(c1);
                                                    paragrafo.Add(NovaLinha);
                                                    paragrafo.Add(c2);
                                                    cellCabec = new PdfPCell(paragrafo);
                                                    //cellCabec.Border = 0;
                                                    if (linhaB1[0].Contains("Razao Social"))
                                                        cellCabec.Colspan = 4;
                                                    cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                    cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                    cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableClientes2.AddCell(cellCabec);
                                                }
                                            }
                                            linhaC1 = linhaA1[6].Split(char.Parse("|")); //  Logradouro, Cidade, UF, Cep, Tel, Fax
                                            foreach (string dados in linhaC1)
                                            {
                                                string[] linhaB2 = dados.Split(char.Parse(":"));
                                                if (linhaB2.Length == 2)
                                                {
                                                    Phrase c1 = new Phrase(linhaB2[0], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL));
                                                    Phrase c2 = new Phrase(linhaB2[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, iTextSharp.text.Font.NORMAL));
                                                    Phrase NovaLinha = new Phrase(Environment.NewLine);

                                                    Paragraph paragrafo = new Paragraph();
                                                    paragrafo.Add(c1);
                                                    paragrafo.Add(NovaLinha);
                                                    paragrafo.Add(c2);
                                                    cellCabec = new PdfPCell(paragrafo);
                                                    //cellCabec.Border = 0;
                                                    if (linhaB2[0].Contains("Logradouro"))
                                                        cellCabec.Colspan = 3;
                                                    cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                    cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                    cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableClientes2.AddCell(cellCabec);
                                                }
                                            }

                                            linhaC1 = linhaA1[7].Split(char.Parse(":"));// Email, Att e Mensagem
                                            if (linhaC1.Length == 2)
                                            {
                                                Phrase c1 = new Phrase(linhaC1[0], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL));
                                                Phrase c2 = new Phrase(linhaC1[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, iTextSharp.text.Font.NORMAL));
                                                Phrase NovaLinha = new Phrase(Environment.NewLine);
                                                Paragraph paragrafo = new Paragraph();
                                                paragrafo.Add(c1);
                                                paragrafo.Add(NovaLinha);
                                                paragrafo.Add(c2);
                                                cellCabec = new PdfPCell(paragrafo);
                                                //cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                tableClientes2.AddCell(cellCabec);

                                                linhaC1 = linhaA1[8].Split(char.Parse(":"));// Email, Att e Mensagem
                                                c1 = new Phrase(linhaC1[0], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL));
                                                c2 = new Phrase(linhaC1[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, iTextSharp.text.Font.NORMAL));
                                                NovaLinha = new Phrase(Environment.NewLine);
                                                paragrafo = new Paragraph();
                                                paragrafo.Add(c1);
                                                paragrafo.Add(NovaLinha);
                                                paragrafo.Add(c2);
                                                cellCabec = new PdfPCell(paragrafo);
                                                //cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                tableClientes2.AddCell(cellCabec);
                                            }
                                            cellCabec = new PdfPCell(new Phrase(linhaA1[9], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 6;
                                            cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                            cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                            tableClientes2.AddCell(cellCabec);

                                            /*-----------------------------------------------------------------------------------------------------------------------------------*/
                                            //TableCabec
                                            cell = new PdfPCell(TableCabec);
                                            cell.Colspan = 15;
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            tableItens.AddCell(cell);

                                            //tableClientes2
                                            cell = new PdfPCell(tableClientes2);
                                            cell.Colspan = 15;
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            tableItens.AddCell(cell);

                                            //table.AddCell(tableItens);
                                            //documento.Add(table);
                                            //documento.NewPage();

                                            //tableItens = new PdfPTable(13);
                                            //tableItens.WidthPercentage = 95;
                                        }

                                        //Imprime os itens
                                        string[] Itens = Detalhes.Split(char.Parse("|"));
                                        if (Itens.Length == 10)
                                        {
                                            AlteraCor = !(AlteraCor);
                                            //seleciona o primeiro item (Cabeçalho)
                                            if (PrimeiroItemCabec == "")
                                            {
                                                PrimeiroItemCabec = Detalhes;
                                            }

                                            //Item
                                            cell = new PdfPCell(new Phrase(Itens[0], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Codigo
                                            cell = new PdfPCell(new Phrase(Itens[1], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Discriminação
                                            cell = new PdfPCell(new Phrase(Itens[2], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                            cell.Colspan = 6;
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Qtde
                                            cell = new PdfPCell(new Phrase(Itens[3], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Class.Fiscal
                                            cell = new PdfPCell(new Phrase(Itens[4], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //% ICMS
                                            cell = new PdfPCell(new Phrase(Itens[5], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Pr.Unit c/ICMS
                                            cell = new PdfPCell(new Phrase(Itens[6], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Pr.Total c/ICMS
                                            string strPrecoTotalComICMS = Itens[7];
                                            cell = new PdfPCell(new Phrase(strPrecoTotalComICMS, FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //IPI
                                            cell = new PdfPCell(new Phrase(Itens[8], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Prazo de Entrega
                                            cell = new PdfPCell(new Phrase(Itens[9], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                                            cell.BackgroundColor = RetornaCor(AlteraCor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);

                                        }
                                        Imprime = true;

                                        break;
                                    case "D":
                                        string[] Totais = Detalhes.Split(char.Parse("|"));
                                        if (LetraAnt != Letra) // IMPRESSO CABEÇALHO DOS TOTAIS
                                        {
                                            if (ContaLinhas < 20) //completa o restante da pagina com linhas em branco
                                            {
                                                int Limite = 23;
                                                //if (PulaPagina) { Limite = 36; }
                                                while (ContaLinhas < Limite)//for (int i = 1; i < (35 - ContaLinhas); i++)
                                                {
                                                    AlteraCor = !(AlteraCor);
                                                    //Item
                                                    cell = new PdfPCell(new Phrase(@"***", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                                    cell.BackgroundColor = RetornaCor(AlteraCor);
                                                    cell.FixedHeight = 14f;
                                                    tableItens.AddCell(cell);
                                                    //Codigo
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                                    cell.BackgroundColor = RetornaCor(AlteraCor);
                                                    cell.FixedHeight = 14f;
                                                    tableItens.AddCell(cell);
                                                    //Discriminação
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                                    cell.Colspan = 6;
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                                    cell.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableItens.AddCell(cell);
                                                    //Qtde
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                    cell.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableItens.AddCell(cell);
                                                    //Class.Fiscal
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                                    cell.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableItens.AddCell(cell);
                                                    //% ICMS
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                    cell.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableItens.AddCell(cell);
                                                    //Pr.Unit c/ICMS
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                    cell.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableItens.AddCell(cell);
                                                    //Pr.Total c/ICMS
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                    cell.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableItens.AddCell(cell);
                                                    //IPI
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                                    cell.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableItens.AddCell(cell);
                                                    //Prazo de Entrega
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, CorFonte)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                                                    cell.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableItens.AddCell(cell);

                                                    ContaLinhas++;
                                                }
                                            }
                                            if (ContaLinhas > 29)
                                            {
                                                AlteraCor = true;
                                                ContaLinhas = 1;
                                                //PulaPagina = true;
                                                ContaPaginas++;
                                                TableCabec = new PdfPTable(6);
                                                TableCabec.WidthPercentage = 100;
                                                tableClientes2 = new PdfPTable(6);
                                                tableClientes2.WidthPercentage = 100;

                                                LogoMarca = Directory.GetCurrentDirectory() + @"\LogoMF2.jpg";
                                                ImgMinas = iTextSharp.text.Image.GetInstance(LogoMarca);
                                                ImgMinas.Alignment = iTextSharp.text.Image.ALIGN_CENTER;


                                                cellCabec = new PdfPCell(ImgMinas);
                                                cellCabec.Rowspan = 4;
                                                cellCabec.Colspan = 2;
                                                cellCabec.Border = 0;
                                                cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                cellCabec.HorizontalAlignment = PdfPCell.ALIGN_MIDDLE;
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(@"vendas@minasferramentas.com.br", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);
                                                string[] linhaC = linhaA1[0].Split(char.Parse(":"));

                                                Chunk c5 = new Chunk(linhaC[0], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 16, iTextSharp.text.Font.NORMAL));
                                                Chunk c6 = new Chunk(linhaC[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 16, iTextSharp.text.Font.NORMAL));
                                                Phrase p2 = new Phrase();
                                                p2.Add(c5);
                                                p2.Add(c6);
                                                cellCabec = new PdfPCell(p2);
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);


                                                cellCabec = new PdfPCell(new Phrase(@"CNPJ.: 17.194.994/0001-27 IE.: 0620080420094", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(linhaA1[1] + @" Pag.: " + ContaPaginas.ToString("##0") + @"/" + TotalPaginas, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(@"Av. Bias Fortes, 1853 - Barro Preto - Belo Horizonte/MG",
                                                                FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(linhaA1[2], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(@"Cep:30.170-012   Fone: (31) 2101-6000 Fax: (31) 2101-6010", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(linhaA1[3], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                TableCabec.AddCell(cellCabec);
                                                /*----------------------------------------------------------------------------------------------------------------------------------*/
                                                string[] linhaC1 = linhaA1[5].Split(char.Parse("|")); // Razão Social, CNPJ, IE
                                                foreach (string dados in linhaC1)
                                                {
                                                    string[] linhaB1 = dados.Split(char.Parse(":"));
                                                    if (linhaB1.Length == 2)
                                                    {
                                                        Chunk c1 = new Chunk(linhaB1[0], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL));
                                                        Chunk c2 = new Chunk(linhaB1[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, iTextSharp.text.Font.NORMAL));
                                                        Phrase NovaLinha = new Phrase(Environment.NewLine);
                                                        Paragraph paragrafo = new Paragraph();
                                                        paragrafo.Add(c1);
                                                        paragrafo.Add(NovaLinha);
                                                        paragrafo.Add(c2);
                                                        cellCabec = new PdfPCell(paragrafo);
                                                        //cellCabec.Border = 0;
                                                        if (linhaB1[0].Contains("Razao Social"))
                                                            cellCabec.Colspan = 4;
                                                        cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                        cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                        cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                        tableClientes2.AddCell(cellCabec);
                                                    }
                                                }
                                                linhaC1 = linhaA1[6].Split(char.Parse("|")); //  Logradouro, Cidade, UF, Cep, Tel, Fax
                                                foreach (string dados in linhaC1)
                                                {
                                                    string[] linhaB2 = dados.Split(char.Parse(":"));
                                                    if (linhaB2.Length == 2)
                                                    {
                                                        Phrase c1 = new Phrase(linhaB2[0], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL));
                                                        Phrase c2 = new Phrase(linhaB2[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, iTextSharp.text.Font.NORMAL));
                                                        Phrase NovaLinha = new Phrase(Environment.NewLine);

                                                        Paragraph paragrafo = new Paragraph();
                                                        paragrafo.Add(c1);
                                                        paragrafo.Add(NovaLinha);
                                                        paragrafo.Add(c2);
                                                        cellCabec = new PdfPCell(paragrafo);
                                                        //cellCabec.Border = 0;
                                                        if (linhaB2[0].Contains("Logradouro"))
                                                            cellCabec.Colspan = 3;
                                                        cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                        cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                        cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                        tableClientes2.AddCell(cellCabec);
                                                    }
                                                }

                                                linhaC1 = linhaA1[7].Split(char.Parse(":"));// Email, Att e Mensagem
                                                if (linhaC1.Length == 2)
                                                {
                                                    Phrase c1 = new Phrase(linhaC1[0], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL));
                                                    Phrase c2 = new Phrase(linhaC1[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, iTextSharp.text.Font.NORMAL));
                                                    Phrase NovaLinha = new Phrase(Environment.NewLine);
                                                    Paragraph paragrafo = new Paragraph();
                                                    paragrafo.Add(c1);
                                                    paragrafo.Add(NovaLinha);
                                                    paragrafo.Add(c2);
                                                    cellCabec = new PdfPCell(paragrafo);
                                                    //cellCabec.Border = 0;
                                                    cellCabec.Colspan = 2;
                                                    cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                    cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                    cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableClientes2.AddCell(cellCabec);

                                                    linhaC1 = linhaA1[8].Split(char.Parse(":"));// Email, Att e Mensagem
                                                    c1 = new Phrase(linhaC1[0], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL));
                                                    c2 = new Phrase(linhaC1[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, iTextSharp.text.Font.NORMAL));
                                                    NovaLinha = new Phrase(Environment.NewLine);
                                                    paragrafo = new Paragraph();
                                                    paragrafo.Add(c1);
                                                    paragrafo.Add(NovaLinha);
                                                    paragrafo.Add(c2);
                                                    cellCabec = new PdfPCell(paragrafo);
                                                    //cellCabec.Border = 0;
                                                    cellCabec.Colspan = 2;
                                                    cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                    cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                    cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                    tableClientes2.AddCell(cellCabec);
                                                }
                                                cellCabec = new PdfPCell(new Phrase(linhaA1[9], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 6;
                                                cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                cellCabec.BackgroundColor = RetornaCor(AlteraCor);
                                                tableClientes2.AddCell(cellCabec);

                                                /*-----------------------------------------------------------------------------------------------------------------------------------*/
                                                //TableCabec
                                                cell = new PdfPCell(TableCabec);
                                                cell.Colspan = 15;
                                                cell.VerticalAlignment = Element.ALIGN_TOP;
                                                cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                                                cell.BackgroundColor = RetornaCor(AlteraCor);
                                                tableItens.AddCell(cell);

                                                //tableClientes2
                                                cell = new PdfPCell(tableClientes2);
                                                cell.Colspan = 15;
                                                cell.VerticalAlignment = Element.ALIGN_TOP;
                                                cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                                                cell.BackgroundColor = RetornaCor(AlteraCor);
                                                tableItens.AddCell(cell);

                                                //table.AddCell(tableItens);
                                                //documento.Add(table);
                                                //documento.NewPage();

                                                //tableItens = new PdfPTable(13);
                                                //tableItens.WidthPercentage = 95;
                                            }

                                            cell = new PdfPCell(new Phrase(Totais[0], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL)));
                                            cell.Rowspan = 2;
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                            tableTotais.AddCell(cell);
                                            //Valor da Mão de Obra
                                            cell = new PdfPCell(new Phrase(@"Valor da Mão de Obra", FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.BOLD)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            tableTotais.AddCell(cell);
                                            //Total do IPI
                                            cell = new PdfPCell(new Phrase(@"Total do IPI", FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.BOLD)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            tableTotais.AddCell(cell);
                                            //Total de Mercadorias
                                            cell = new PdfPCell(new Phrase(@"Total de Mercadorias", FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.BOLD)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            tableTotais.AddCell(cell);
                                            //Total Geral
                                            cell = new PdfPCell(new Phrase(@"Total Geral", FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.BOLD)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            tableTotais.AddCell(cell);
                                            LetraAnt = Letra;
                                        }
                                        if (Totais.Length == 5)
                                        {
                                            //Valor da Mão de Obra
                                            cell = new PdfPCell(new Phrase(Totais[1].Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.BOLD, BaseColor.BLUE)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            tableTotais.AddCell(cell);
                                            //Total do IPI
                                            cell = new PdfPCell(new Phrase(Totais[2].Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.BOLD, BaseColor.BLUE)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            tableTotais.AddCell(cell);
                                            //Total de Mercadorias
                                            cell = new PdfPCell(new Phrase(Totais[3].Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.BOLD, BaseColor.BLUE)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            tableTotais.AddCell(cell);
                                            //Total Geral
                                            cell = new PdfPCell(new Phrase(Totais[4].Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.BOLD, BaseColor.BLUE)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            tableTotais.AddCell(cell);
                                        }
                                        table.AddCell(tableTotais);
                                        break;
                                    case "F":
                                        if (LetraAnt != Letra) // INVERTE A LETRA
                                        {
                                            LetraAnt = Letra;
                                        }
                                        string[] Obs = Detalhes.Split(char.Parse(":"));
                                        if (Obs.Length == 2)
                                        {
                                            if (!Obs[0].Contains("OBS"))
                                            {
                                                Phrase c1 = new Phrase(Obs[0].Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL));
                                                Phrase c2 = new Phrase(Obs[1].Trim(), FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, iTextSharp.text.Font.BOLD));
                                                Phrase NovaLinha = new Phrase(Environment.NewLine);
                                                Paragraph paragrafo = new Paragraph();
                                                paragrafo.Add(c1);
                                                paragrafo.Add(NovaLinha);
                                                paragrafo.Add(c2);
                                                cell = new PdfPCell(paragrafo);
                                                //cell.Border = 1;
                                                cell.VerticalAlignment = Element.ALIGN_TOP;
                                                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                                tableObs.AddCell(cell);
                                            }
                                            else
                                            {
                                                cell = new PdfPCell(new Phrase(Obs[1].Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL)));
                                                cell.Colspan = 6;
                                                cell.Border = 1;
                                                tableObs.AddCell(cell);
                                            }
                                        }

                                        break;
                                    case "G":
                                        if (LetraAnt != Letra) // IMPRESSO AS OBSERVACOES
                                        {
                                            cell = new PdfPCell(tableObs);
                                            cell.Border = 0;
                                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                            table.AddCell(cell);

                                            //cell = new PdfPCell(new Phrase(Detalhes.Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL)));
                                            //cell.Border = 0;
                                            //cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                            //table.AddCell(cell);

                                            intLinha = 1;
                                            LetraAnt = Letra;

                                        }

                                        switch (intLinha)
                                        {
                                            case 1: cell = new PdfPCell(new Phrase(Detalhes.Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL)));
                                                cell.Border = 0;
                                                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                                tableObs2.AddCell(cell);
                                                cell = new PdfPCell(new Phrase(@"DISTRIBUIDOR AUTORIZADO de ferramentas de metal duro  LAMINA, Consulte-nos", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL)));
                                                cell.Border = 0;
                                                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                                tableObs2.AddCell(cell);
                                                break;
                                            case 2: cell = new PdfPCell(new Phrase(Detalhes.Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL)));
                                                cell.Border = 0;
                                                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                                tableObs2.AddCell(cell);
                                                cell = new PdfPCell(new Phrase(@"ESPECIALIZADA EM FERRAMENTAS PARA INDUSTRIA E MECANICA EM GERAL", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL)));
                                                cell.Border = 0;
                                                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                                tableObs2.AddCell(cell);
                                                break;
                                            default:
                                                cell = new PdfPCell(new Phrase(Detalhes.Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL)));
                                                cell.Border = 0;
                                                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                                tableObs2.AddCell(cell);
                                                cell = new PdfPCell(new Phrase(@" ", FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL)));
                                                cell.Border = 0;
                                                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                                tableObs2.AddCell(cell);
                                                break;

                                        }
                                        intLinha++;
                                        break;
                                    case "H":
                                        if (LetraAnt != Letra) // IMPRESSO AS OBSERVACOES
                                        {
                                            intLinha = 1;
                                            LetraAnt = Letra;
                                        }
                                        switch (intLinha)
                                        {
                                            case 1: Observacao1 = Detalhes.Trim(); break;
                                            case 2: Observacao2 = Detalhes.Trim(); break;
                                        }
                                        intLinha++;
                                        break;
                                    //default:
                                    //    cell = new PdfPCell(new Phrase(Detalhes.Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL)));
                                    //    cell.Border = 0;
                                    //    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                    //    table.AddCell(cell);
                                    //    break;


                                }
                            }
                        }
                    }
                }
                cell = new PdfPCell(tableObs2);
                cell.Border = 0;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                table.AddCell(cell);

                documento.Add(table);
                srArquivo.Close();
                // fecha o documento
                Writer.Flush();
                documento.Close();

                MovePara(PrimeiraLetra, ColetaTXT, CodVendedor, ColetaPDFAenviar);

            }
            catch (IOException e)
            {
                Console.WriteLine(ColetaTXT + " Erro->" + e.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(ColetaTXT + " Erro->" + e.Message);
            }

            //Console.WriteLine("Documento PDF criado com sucesso.");
        }

        //============= gera PDF para enviar por e-mail ============================================================================================
        private static void MovePara(string OndeGerar, string ColetaTXT, int CodVendedor, string ColetaPDFAenviar)
        {
            string enviaMensagem = "";
            int statusEmail = 0;
            string Motivo = "";
            try
            {
                string sourceFile = CaminhoOrigem + @"\" + ColetaPDFAenviar;
                Cotacao = Cotacao.Replace(".", "").Replace(":", "");
                //string destFile = @"F:\Aenviar\" + CnpjCpf.Replace(".", "").Replace("-", "").Replace("/", "").Replace(" ", "") +
                //    @"\" + DateTime.Now.Year + "-" + DateTime.Now.Month.ToString("00") + "-" + DateTime.Now.Day.ToString("00") + "-"
                //    + DateTime.Now.Hour.ToString("00") + "-" + DateTime.Now.Minute.ToString("00") + "-" + DateTime.Now.Second.ToString("00") + "-"
                //    + Cotacao + "-" + CodVendedor.ToString("000") + ".pdf";

                string destFile = @"F:\Aenviar\" + CnpjCpf.Replace(".", "").Replace("-", "").Replace("/", "").Replace(" ", "") +
                    @"\" + Cotacao + "-" + DateTime.Now.Hour.ToString("00") + "-" +
                    DateTime.Now.Minute.ToString("00") + "-" + DateTime.Now.Second.ToString("00") + ".pdf";

                string existeCaminhoDestino = @"F:\Aenviar\" + CnpjCpf.Replace(".", "").Replace("-", "").Replace("/", "").Replace(" ", "");
                if (!System.IO.Directory.Exists(existeCaminhoDestino))
                {
                    System.IO.Directory.CreateDirectory(existeCaminhoDestino);
                }

                //============== imprime na impressora 1 (Recepção) ======================================================================================
                if (OndeGerar == "C" || OndeGerar == "Q" || OndeGerar == "Y")
                {
                    if (System.IO.File.Exists(destFile)) { System.IO.File.Delete(destFile); }
                    try
                    {
                        System.IO.File.Move(sourceFile, destFile);
                        Console.WriteLine();
                        Console.WriteLine("Movido o PDF para a pasta do Cliente {0} ", CnpjCpf);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(sourceFile + " ++++> " + e.Message);
                        return;
                    }
                    System.Threading.Thread.Sleep(1500);
                    Console.WriteLine("Aguarde imprimindo coleta...{0}", ColetaPDFAenviar);
                    Process proc = new Process();
                    proc.StartInfo.CreateNoWindow = false;
                    proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    proc.StartInfo.Verb = "print";
                    proc.StartInfo.FileName = destFile;
                    proc.Start();
                    proc.WaitForInputIdle();
                    //proc.WaitForExit(10000);
                    proc.CloseMainWindow();
                    proc.Close();
                    Console.WriteLine();
                    Console.WriteLine("Enviado o Arquivo {0} para impressora.", ColetaPDFAenviar);
                }
                else
                {
                    //================== move para o aenviar a coleta do vendedor =================================================================================
                    //System.Threading.Thread.Sleep(2000);
                    if (System.IO.File.Exists(destFile)) { System.IO.File.Delete(destFile); }
                    try
                    {
                        System.IO.File.Move(sourceFile, destFile);
                        Console.WriteLine();
                        Console.WriteLine("Movido o PDF para a pasta do Cliente {0} ", CnpjCpf);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(sourceFile + " ++++> " + e.Message);
                        return;
                    }

                    //destFile = @"F:\Aenviar\" + CodVendedor + @"\" + ColetaPDFAenviar;

                    if (IsNumeric(Cliente.Trim()))
                    {
                        Cliente = EmailCli;
                    }

                    ContaEmail = RetornaContaDeContato(CodVendedor);
                    if (ValidaEnderecoEmail(EmailCli))
                    {

                        if (Comprador.Trim() == "") { Comprador = "Comprador"; }
                        if (EmailCli.Trim() == "") { EmailCli = "vendas@minasferramentas.com.br"; }
                        if (Observacao1.Trim() != "")
                        {
                            Observacao1 = "<b>Ps:</b> " + Observacao1.Trim() + "<br />" + Observacao2.Trim() + "<br /><br />";
                        }
                        string Assunto = Cotacao + " " + Cliente;
                        enviaMensagem = CorpoEmail + "<br /><br />" +
                                Observacao1 +
                                "Atenciosamente," + "<br />" +
                                  Vendedor.Substring(5, Vendedor.Length - 5) + "<br />" +
                                "Departamento de Vendas." + "<br />" +
                                ContaEmail + "<br />" +
                                "Tel.: (31) 2101.6000  / Fax: (31) 2101.6010<br />" +
                                "Av. Bias Fortes, 1853 | B. Barro Preto | Belo Horizonte - MG | Cep 30170-012 ";

                        // cria uma mensagem para o cliente
                        MailMessage mensagemEmail = new MailMessage(); //(ContaEmail, EmailCli, Assunto, enviaMensagem);
                        mensagemEmail.Sender = new MailAddress(ContaEmail, "Minas Ferramentas - " + Vendedor.Substring(5, Vendedor.Length - 5));
                        mensagemEmail.From = new MailAddress(ContaEmail, "Minas Ferramentas - " + Vendedor.Substring(5, Vendedor.Length - 5));
                        mensagemEmail.To.Add(new MailAddress(EmailCli));
                        mensagemEmail.Subject = Assunto;
                        mensagemEmail.Body = enviaMensagem;
                        mensagemEmail.IsBodyHtml = true;
                        mensagemEmail.Priority = MailPriority.Normal;
                        Attachment anexo = new Attachment(destFile);
                        mensagemEmail.Attachments.Add(anexo);
                        SmtpClient client = new SmtpClient();
                        client.Host = ServidorSMTP;
                        client.Port = 587;
                        client.EnableSsl = false;
                        client.UseDefaultCredentials = false;
                        client.Credentials = new NetworkCredential(ContaEmail, Senha);
                        // envia a mensagem
                        client.Send(mensagemEmail);


                        // envia uma copia para a conta copiadeemail@minasferramentas.com.br
                        mensagemEmail = new MailMessage(); //(ContaEmail, EmailCli, Assunto, enviaMensagem);
                        mensagemEmail.Sender = new MailAddress(ContaEmail, "Minas Ferramentas - " + Vendedor.Substring(5, Vendedor.Length - 5));
                        mensagemEmail.From = new MailAddress(ContaEmail, "Minas Ferramentas - " + Vendedor.Substring(5, Vendedor.Length - 5));
                        mensagemEmail.To.Add(new MailAddress("copiadeemail@minasferramentas.com.br", Cliente));
                        enviaMensagem = Cotacao + " " + Cliente + "<br />E-mail:" + EmailCli + "<br /><hr><br />" + enviaMensagem;
                        mensagemEmail.Subject = Assunto;
                        mensagemEmail.Body = enviaMensagem;
                        mensagemEmail.IsBodyHtml = true;
                        mensagemEmail.Priority = MailPriority.High;
                        anexo = new Attachment(destFile);
                        mensagemEmail.Attachments.Add(anexo);
                        client = new SmtpClient();
                        client.Host = ServidorSMTP;
                        client.Port = 587;
                        client.EnableSsl = false;
                        client.UseDefaultCredentials = false;
                        client.Credentials = new NetworkCredential(ContaEmail, Senha);
                        // envia a mensagem
                        client.Send(mensagemEmail);


                        // envia confirmacao para o vendedor de mensagem enviada OK
                        Assunto = "OK - " + Cotacao + " " + Cliente + " " + DateTime.Now.ToString();
                        enviaMensagem = Cotacao + "<br />" + CnpjCpf + "-" + Cliente
                            + "<br /><b>Enviada com sucesso para:</b><br />" + EmailCli
                             + "<br />Em: " + DateTime.Now.ToString() + ".<br />" +
                             Observacao1 + ".";

                        mensagemEmail = new MailMessage(); //(ContaEmail, EmailCli, Assunto, enviaMensagem);
                        mensagemEmail.Sender = new MailAddress("vendas@minasferramentas.com.br", "Minas Ferramentas Ltda");
                        mensagemEmail.From = new MailAddress("vendas@minasferramentas.com.br", "Minas Ferramentas Ltda");
                        mensagemEmail.To.Add(new MailAddress(ContaEmail, Vendedor));
                        mensagemEmail.Subject = Assunto;
                        mensagemEmail.Body = enviaMensagem;
                        mensagemEmail.IsBodyHtml = true;
                        mensagemEmail.Priority = MailPriority.High;
                        client = new SmtpClient();
                        client.Host = ServidorSMTP;
                        client.Port = 587;
                        client.EnableSsl = false;
                        client.UseDefaultCredentials = false;
                        client.Credentials = new NetworkCredential("vendas@minasferramentas.com.br", "M1n4aSf*");
                        // envia a mensagem
                        client.Send(mensagemEmail);

                        statusEmail = 1;
                        Console.WriteLine();
                        Console.WriteLine(Assunto);
                    }
                    else
                    {
                        string Assunto = "((ERRO)) - " + Cotacao + " " + Cliente + " " + DateTime.Now.ToString();
                        enviaMensagem = Cotacao + "<br />" + CnpjCpf + "-" + Cliente + " - Vendedor: " + Vendedor
                        + "<br /><b>NAO enviada para:</b><br />" + EmailCli + "<br />Motivo: <b>Email Invalido.</b>" +
                        "<br />Em: " + DateTime.Now.ToString() + ".";
                        MailMessage mensagemEmail = new MailMessage(); //(ContaEmail, EmailCli, Assunto, enviaMensagem);
                        mensagemEmail.Sender = new MailAddress("vendas@minasferramentas.com.br", "Minas Ferramentas Ltda");
                        mensagemEmail.From = new MailAddress("vendas@minasferramentas.com.br", "Minas Ferramentas Ltda");
                        mensagemEmail.To.Add(new MailAddress(ContaEmail, Vendedor));
                        mensagemEmail.Subject = Assunto;
                        mensagemEmail.Body = enviaMensagem;
                        mensagemEmail.IsBodyHtml = true;
                        mensagemEmail.Priority = MailPriority.High;
                        SmtpClient client = new SmtpClient();
                        client.Host = ServidorSMTP;
                        client.Port = 587;
                        client.EnableSsl = false;
                        client.UseDefaultCredentials = false;
                        client.Credentials = new NetworkCredential("vendas@minasferramentas.com.br", "M1n4aSf*");
                        // envia a mensagem
                        client.Send(mensagemEmail);
                        statusEmail = 2;
                        Motivo = "Email Invalido.";
                        Console.WriteLine();
                        Console.WriteLine(Assunto);
                    }
                }
            }
            catch (Exception e)
            {
                string Assunto = "((ERRO)) - " + Cotacao + " " + Cliente + " " + DateTime.Now.ToString();
                enviaMensagem = Cotacao + "<br />" + CnpjCpf + "-" + Cliente + " - Vendedor: " + Vendedor
                        + "<br /><b>NAO enviada para:</b><br />" + EmailCli + "<br />Motivo: " + e.Message +
                        "<br />Em: " + DateTime.Now.ToString() + ".";
                MailMessage mensagemEmail = new MailMessage(); //(ContaEmail, EmailCli, Assunto, enviaMensagem);
                mensagemEmail.Sender = new MailAddress("vendas@minasferramentas.com.br", "Minas Ferramentas Ltda");
                mensagemEmail.From = new MailAddress("vendas@minasferramentas.com.br", "Minas Ferramentas Ltda");
                mensagemEmail.To.Add(new MailAddress(ContaEmail));
                mensagemEmail.Bcc.Add(new MailAddress("copiadeemail@minasferramentas.com.br"));

                mensagemEmail.Subject = Assunto;
                mensagemEmail.Body = enviaMensagem;
                mensagemEmail.IsBodyHtml = true;
                mensagemEmail.Priority = MailPriority.High;
                SmtpClient client = new SmtpClient();
                client.Host = ServidorSMTP;
                client.Port = 587;
                client.EnableSsl = false;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential("vendas@minasferramentas.com.br", "M1n4aSf*");
                // envia a mensagem
                client.Send(mensagemEmail);
                Console.WriteLine();
                Console.WriteLine(Assunto);
                statusEmail = 2;
                Motivo = e.Message;
            }
            finally
            {
                // ========= copia o arquivo txt para a pasta impressos
                string sourceFile = System.IO.Path.Combine(CaminhoOrigem, ColetaTXT);
                string destFile = System.IO.Path.Combine(CaminhoDestino, ColetaTXT);
                try
                {
                    if (!System.IO.Directory.Exists(CaminhoDestino))
                    {
                        System.IO.Directory.CreateDirectory(CaminhoDestino);
                    }
                    if (System.IO.File.Exists(destFile))
                    {
                        System.IO.File.Delete(destFile);
                    }
                    //System.Threading.Thread.Sleep(1500);
                    System.IO.File.Move(sourceFile, destFile);
                    Console.WriteLine("Movido arquivo de {0} para {1}", sourceFile, destFile);
                    //System.IO.File.Delete(sourceFile);
                    //Console.WriteLine("Excluido arquivo " + sourceFile);
                }
                catch (Exception e)
                {
                    Console.WriteLine(sourceFile + " >>> " + e.Message);
                }

                if (statusEmail != 0)
                {
                    //gera txt com os dados do envio.
                    StreamWriter s = File.AppendText(CaminhoDestino + @"email.txt");
                    EmailCli = CompletaCampos(EmailCli, 50);
                    CnpjCpf = (CnpjCpf.Replace(".", "").Replace("-", "").Replace("/", "").Replace(" ", "") + "00").Trim();
                    if (CnpjCpf.Length == 13)
                    {
                        CnpjCpf = CnpjCpf + "000";
                    }

                    string linha = "|" + CodVendedor.ToString("0000") + CnpjCpf + Cotacao.Substring(11, 4) + EmailCli +
                            DateTime.Now.ToString("ddMMyy") + statusEmail + DateTime.Now.ToString("HHmmss") + "00";

                    if (statusEmail == 1)
                        linha = linha + ("Enviada com sucesso para: " + EmailCli.Trim() + " " +
                            CnpjCpf + "-" + Cliente.Trim() + " Em: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") +
                            ".=================================================").Substring(0, 100);

                    else
                        linha = linha + ("Motivo: " + Motivo + " " +
                        CnpjCpf + "-" + Cliente.Trim() + " Em: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") +
                        ".=========================================================================================================="+
                        "=================================================================================================").Substring(0, 100);
                    s.WriteLine(linha);
                    s.Close();
                }
            }
        }

        private static bool IsNumeric(string data)
        {
            bool isnumeric = false;
            char[] datachars = data.ToCharArray();

            foreach (var datachar in datachars)
                isnumeric = isnumeric ? char.IsDigit(datachar) : isnumeric;

            return isnumeric;
        }

        private static int ContadorPaginas(string caminho)
        {
            StreamReader srArquivo = new StreamReader(caminho);
            int ContaLinhas = 0;
            int ContaPaginas = 0;
            while (srArquivo.Peek() != -1)
            {
                string strLinha = srArquivo.ReadLine();
                if (strLinha != "")
                {
                    int TamanhoLinha = strLinha.Length;
                    string Letra = strLinha.Substring(0, 1);
                    if (TamanhoLinha > 1)
                    {
                        switch (Letra)
                        {
                            case "C": ContaLinhas++;
                                break;
                        }
                    }
                }
            }
            if (ContaLinhas > 28)
            {
                ContaPaginas = (int)ContaLinhas / 28;
                if ((ContaLinhas % 28) != 0) { ContaPaginas++; }
            }
            else
                ContaPaginas = 1;

            return ContaPaginas;
        }

        private static BaseColor RetornaCor(bool AlteraCor)
        {
            if (AlteraCor)
            {
                //CorFonte = BaseColor.WHITE;
                CorFonte = BaseColor.BLACK;
                return BaseColor.WHITE;
            }
            else
            {
                CorFonte = BaseColor.BLACK;
                return BaseColor.LIGHT_GRAY;
            }
        }

        private static bool ValidaEnderecoEmail(string enderecoEmail)
        {
            try
            {
                //define a expressão regulara para validar o email
                string texto_Validar = enderecoEmail;
                Regex expressaoRegex = new Regex(@"\w+@[a-zA-Z_0-9-]+?\.[a-zA-Z]{2,3}");

                // testa o email com a expressão
                if (expressaoRegex.IsMatch(texto_Validar))
                {
                    // o email é valido
                    return true;
                }
                else
                {
                    // o email é inválido
                    return false;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private static string CompletaCampos(string campo, int tamanho)
        {
            int resto = tamanho - campo.Length;
            while (resto != 0)
            {
                campo += " ";
                resto--;
            }
            return campo;
        }

        private static string RetornaContaDeContato(int CodVendedor)
        {
            string contaEmail = "";
            string CaminhoXML = Directory.GetCurrentDirectory() + @"\ListaEmails.xml";
            DataSet Ds = new DataSet();
            Ds.ReadXml(CaminhoXML);
            DataTable Dt = Ds.Tables[0];
            foreach (DataRow Dr in Dt.Rows)
            {
                int codigo = int.Parse(Dr["codigo"].ToString());
                if (codigo == CodVendedor)
                {
                    string nome = Dr["nome"].ToString();
                    contaEmail = Dr["email"].ToString();
                    Senha = Dr["senha"].ToString();
                }
            }
            if (contaEmail == "")
            {
                contaEmail = "vendas@minasferramentas.com.br";
                Senha = "M1n4aSf*";
            }
            return contaEmail;
        }
    }
}
