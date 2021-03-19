using GeracaoColetasPDF.Entities;
using iTextSharp.text;
using System;
using System.IO;
using System.Diagnostics;
using System.Net.Mail;
using System.Net;
using System.Data;
using System.Text.RegularExpressions;
using iTextSharp.text.pdf;

namespace GeracaoColetasPDF.Services
{
    class GeraPDF
    {
        public string FirstLettre { get; set; }
        public string Printer1 { get; set; }
        public string Printer2 { get; set; }
        public bool Printer { get; set; }
        public string SalesPerson { get; set; }
        public string Buyer { get; set; }
        public string  Customer { get; set; }
        public string NumOrder { get; set; }
        public string CnpjCpf { get; set; }
        public string EmailCustomer { get; set; }
        public string  EmailBody { get; set; }
        public string Note1 { get; set; }
        public string Note2 { get; set; }
        public BaseColor FontColor { get; set; }
        public bool AlternateColor { get; set; }
        public Config Config { get; set; } = new Config();

        public void GerarPDF(string fileTxt)
        {
            Document documento = new Document(PageSize.A4, 5, 5, 20, 20);
            documento.AddAuthor("Minas Ferramentas Ltda");
            documento.AddSubject("Cotação " + fileTxt.Replace(".txt", "").Replace(".TXT", ""));
            documento.SetPageSize(PageSize.A4.Rotate());
            int sallerNumber = 0;

            Config.GetConfig();

            DateTime Dt1 = DateTime.Now;
            string ColetaPDF = "MFL Cotação nº " + fileTxt.Replace(".txt", "").Replace(".TXT", "");
            ColetaPDF = ColetaPDF + " em " + Dt1.ToString("O") + ".pdf";
            string ColetaPDFAenviar = ColetaPDF;
            if (File.Exists(Config.SourcePath + ColetaPDF))
            {
                File.Delete(ColetaPDF);
            }
            ColetaPDF = Config.SourcePath + ColetaPDF;
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
                AlternateColor = true;
                FontColor = BaseColor.BLACK;
                SalesPerson = "";
                Buyer = "";
                Customer = "";
                NumOrder = "";
                CnpjCpf = "";
                EmailCustomer = "";
                Note1 = "";
                Note2 = "";
                int TotalPaginas = CountPages(Config.SourcePath + fileTxt);
                int ContaPaginas = 1;
                StreamReader srArquivo = new StreamReader(Config.SourcePath + fileTxt);
                while (!srArquivo.EndOfStream)
                {
                    string strLinha = srArquivo.ReadLine();
                    if (strLinha != "")
                    {
                        int TamanhoLinha = strLinha.Length;
                        string Letra = strLinha.Substring(0, 1);
                        if ((TamanhoLinha == 4) && (Linha1))
                        {
                            FirstLettre = strLinha.Substring(0, 1);
                            sallerNumber = int.Parse(strLinha.Substring(1, 3));
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

                                                Chunk c5 = new Chunk(linhaA[0], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 16, iTextSharp.text.Font.NORMAL));
                                                Chunk c6 = new Chunk(linhaA[1], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 16, iTextSharp.text.Font.NORMAL));
                                                Phrase p2 = new Phrase();
                                                p2.Add(c5);
                                                p2.Add(c6);
                                                cellCabec = new PdfPCell(p2);
                                                AddCell(ref TableCabec, ref cellCabec, 0, 2, Element.ALIGN_MIDDLE, Element.ALIGN_CENTER, GetColor(AlternateColor));

                                                linhaA1[0] = Detalhes;
                                                NumOrder = Detalhes;
                                                break;
                                            case 1://Data da cotacao e pagina
                                                cellCabec = new PdfPCell(new Phrase(Detalhes + @" Pag.: " + ContaPaginas.ToString("##0") + @"/" + TotalPaginas, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);
                                                AddCell(ref TableCabec, ref cellCabec, 0, 2, Element.ALIGN_MIDDLE, Element.ALIGN_CENTER, GetColor(AlternateColor));

                                                linhaA1[1] = Detalhes;
                                                break;
                                            case 2://Vendedor
                                                cellCabec = new PdfPCell(new Phrase(Detalhes, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);
                                                SalesPerson = linhaA[1];
                                                linhaA1[2] = Detalhes;
                                                break;
                                            case 3://Referente
                                                cellCabec = new PdfPCell(new Phrase(Detalhes, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);
                                                linhaA1[3] = Detalhes;
                                                break;
                                            case 4:
                                                cellCabec = new PdfPCell(new Phrase(Detalhes, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.BOLDITALIC)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 6;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);
                                                linhaA1[4] = Detalhes;
                                                break;

                                            default:
                                                cellCabec = new PdfPCell(new Phrase(@"Belo Horizonte-MG CEP 30.170-012", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(@"Tel.: (31) 2101.6000 - Fax: (31) 2101.6010", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                                Customer = linhaB1[1];
                                                            }
                                                            if (linhaB1[0].Contains("CNPJ"))
                                                                CnpjCpf = linhaB1[1];
                                                            cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                            cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                            cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                            cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                            EmailCustomer = linhaB3[1];
                                                        if (linhaB3[0].Contains("ATT"))
                                                            Buyer = linhaB3[1];

                                                        //cellCabec.Border = 0;
                                                        cellCabec.Colspan = 2;
                                                        cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                        cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                        cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                        cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                    cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                            cell.BackgroundColor = GetColor(AlternateColor);
                                            tableItens.AddCell(cell);

                                            //tableClientes2
                                            cell = new PdfPCell(tableClientes2);
                                            cell.Colspan = 15;
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                                            cell.BackgroundColor = GetColor(AlternateColor);
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
                                            AlternateColor = true;
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
                                            cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                            cellCabec.BackgroundColor = GetColor(AlternateColor);
                                            TableCabec.AddCell(cellCabec);


                                            cellCabec = new PdfPCell(new Phrase(@"CNPJ.: 17.194.994/0001-27 IE.: 0620080420094", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cellCabec.BackgroundColor = GetColor(AlternateColor);
                                            TableCabec.AddCell(cellCabec);

                                            cellCabec = new PdfPCell(new Phrase(linhaA1[1] + @" Pag.: " + ContaPaginas.ToString("##0") + @"/" + TotalPaginas, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cellCabec.BackgroundColor = GetColor(AlternateColor);
                                            TableCabec.AddCell(cellCabec);

                                            cellCabec = new PdfPCell(new Phrase(@"Av. Bias Fortes, 1853 - Barro Preto - Belo Horizonte/MG",
                                                            FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cellCabec.BackgroundColor = GetColor(AlternateColor);
                                            TableCabec.AddCell(cellCabec);

                                            cellCabec = new PdfPCell(new Phrase(linhaA1[2], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cellCabec.BackgroundColor = GetColor(AlternateColor);
                                            TableCabec.AddCell(cellCabec);

                                            cellCabec = new PdfPCell(new Phrase(@"Cep:30.170-012   Fone: (31) 2101-6000 Fax: (31) 2101-6010", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cellCabec.BackgroundColor = GetColor(AlternateColor);
                                            TableCabec.AddCell(cellCabec);

                                            cellCabec = new PdfPCell(new Phrase(linhaA1[3], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 2;
                                            cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                            cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                    cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                    cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                tableClientes2.AddCell(cellCabec);
                                            }
                                            cellCabec = new PdfPCell(new Phrase(linhaA1[9], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL)));
                                            cellCabec.Border = 0;
                                            cellCabec.Colspan = 6;
                                            cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                            cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                            cellCabec.BackgroundColor = GetColor(AlternateColor);
                                            tableClientes2.AddCell(cellCabec);

                                            /*-----------------------------------------------------------------------------------------------------------------------------------*/
                                            //TableCabec
                                            cell = new PdfPCell(TableCabec);
                                            cell.Colspan = 15;
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                                            cell.BackgroundColor = GetColor(AlternateColor);
                                            tableItens.AddCell(cell);

                                            //tableClientes2
                                            cell = new PdfPCell(tableClientes2);
                                            cell.Colspan = 15;
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                                            cell.BackgroundColor = GetColor(AlternateColor);
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
                                            AlternateColor = !(AlternateColor);
                                            //seleciona o primeiro item (Cabeçalho)
                                            if (PrimeiroItemCabec == "")
                                            {
                                                PrimeiroItemCabec = Detalhes;
                                            }

                                            //Item
                                            cell = new PdfPCell(new Phrase(Itens[0], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BackgroundColor = GetColor(AlternateColor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Codigo
                                            cell = new PdfPCell(new Phrase(Itens[1], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BackgroundColor = GetColor(AlternateColor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Discriminação
                                            cell = new PdfPCell(new Phrase(Itens[2], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                            cell.Colspan = 6;
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                            cell.BackgroundColor = GetColor(AlternateColor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Qtde
                                            cell = new PdfPCell(new Phrase(Itens[3], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BackgroundColor = GetColor(AlternateColor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Class.Fiscal
                                            cell = new PdfPCell(new Phrase(Itens[4], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BackgroundColor = GetColor(AlternateColor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //% ICMS
                                            cell = new PdfPCell(new Phrase(Itens[5], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BackgroundColor = GetColor(AlternateColor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Pr.Unit c/ICMS
                                            cell = new PdfPCell(new Phrase(Itens[6], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BackgroundColor = GetColor(AlternateColor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Pr.Total c/ICMS
                                            string strPrecoTotalComICMS = Itens[7];
                                            cell = new PdfPCell(new Phrase(strPrecoTotalComICMS, FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            cell.BackgroundColor = GetColor(AlternateColor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //IPI
                                            cell = new PdfPCell(new Phrase(Itens[8], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                            cell.BackgroundColor = GetColor(AlternateColor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);
                                            //Prazo de Entrega
                                            cell = new PdfPCell(new Phrase(Itens[9], FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                            cell.VerticalAlignment = Element.ALIGN_TOP;
                                            cell.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                                            cell.BackgroundColor = GetColor(AlternateColor);
                                            cell.FixedHeight = 14f;
                                            tableItens.AddCell(cell);

                                        }
                                        Print = true;

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
                                                    AlternateColor = !(AlternateColor);
                                                    //Item
                                                    cell = new PdfPCell(new Phrase(@"***", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                                    cell.BackgroundColor = GetColor(AlternateColor);
                                                    cell.FixedHeight = 14f;
                                                    tableItens.AddCell(cell);
                                                    //Codigo
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                                    cell.BackgroundColor = GetColor(AlternateColor);
                                                    cell.FixedHeight = 14f;
                                                    tableItens.AddCell(cell);
                                                    //Discriminação
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                                    cell.Colspan = 6;
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                                    cell.BackgroundColor = GetColor(AlternateColor);
                                                    tableItens.AddCell(cell);
                                                    //Qtde
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                    cell.BackgroundColor = GetColor(AlternateColor);
                                                    tableItens.AddCell(cell);
                                                    //Class.Fiscal
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                                    cell.BackgroundColor = GetColor(AlternateColor);
                                                    tableItens.AddCell(cell);
                                                    //% ICMS
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                    cell.BackgroundColor = GetColor(AlternateColor);
                                                    tableItens.AddCell(cell);
                                                    //Pr.Unit c/ICMS
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                    cell.BackgroundColor = GetColor(AlternateColor);
                                                    tableItens.AddCell(cell);
                                                    //Pr.Total c/ICMS
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                    cell.BackgroundColor = GetColor(AlternateColor);
                                                    tableItens.AddCell(cell);
                                                    //IPI
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                                    cell.BackgroundColor = GetColor(AlternateColor);
                                                    tableItens.AddCell(cell);
                                                    //Prazo de Entrega
                                                    cell = new PdfPCell(new Phrase(@"", FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL, FontColor)));
                                                    cell.VerticalAlignment = Element.ALIGN_TOP;
                                                    cell.HorizontalAlignment = Element.ALIGN_JUSTIFIED;
                                                    cell.BackgroundColor = GetColor(AlternateColor);
                                                    tableItens.AddCell(cell);

                                                    ContaLinhas++;
                                                }
                                            }
                                            if (ContaLinhas > 29)
                                            {
                                                AlternateColor = true;
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
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);


                                                cellCabec = new PdfPCell(new Phrase(@"CNPJ.: 17.194.994/0001-27 IE.: 0620080420094", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(linhaA1[1] + @" Pag.: " + ContaPaginas.ToString("##0") + @"/" + TotalPaginas, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(@"Av. Bias Fortes, 1853 - Barro Preto - Belo Horizonte/MG",
                                                                FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(linhaA1[2], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(@"Cep:30.170-012   Fone: (31) 2101-6000 Fax: (31) 2101-6010", FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_CENTER;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                TableCabec.AddCell(cellCabec);

                                                cellCabec = new PdfPCell(new Phrase(linhaA1[3], FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 2;
                                                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                        cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                        cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                    cellCabec.BackgroundColor = GetColor(AlternateColor);
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
                                                    cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                    tableClientes2.AddCell(cellCabec);
                                                }
                                                cellCabec = new PdfPCell(new Phrase(linhaA1[9], FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.NORMAL)));
                                                cellCabec.Border = 0;
                                                cellCabec.Colspan = 6;
                                                cellCabec.VerticalAlignment = Element.ALIGN_TOP;
                                                cellCabec.HorizontalAlignment = Element.ALIGN_LEFT;
                                                cellCabec.BackgroundColor = GetColor(AlternateColor);
                                                tableClientes2.AddCell(cellCabec);

                                                /*-----------------------------------------------------------------------------------------------------------------------------------*/
                                                //TableCabec
                                                cell = new PdfPCell(TableCabec);
                                                cell.Colspan = 15;
                                                cell.VerticalAlignment = Element.ALIGN_TOP;
                                                cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                                                cell.BackgroundColor = GetColor(AlternateColor);
                                                tableItens.AddCell(cell);

                                                //tableClientes2
                                                cell = new PdfPCell(tableClientes2);
                                                cell.Colspan = 15;
                                                cell.VerticalAlignment = Element.ALIGN_TOP;
                                                cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                                                cell.BackgroundColor = GetColor(AlternateColor);
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
                                            case 1:
                                                cell = new PdfPCell(new Phrase(Detalhes.Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL)));
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
                                            case 2:
                                                cell = new PdfPCell(new Phrase(Detalhes.Trim(), FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL)));
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
                                            case 1: Note1 = Detalhes.Trim(); break;
                                            case 2: Note2 = Detalhes.Trim(); break;
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

                MovePara(FirstLettre, fileTxt, sallerNumber, ColetaPDFAenviar);

            }
            catch (IOException e)
            {
                Console.WriteLine(fileTxt + " Erro->" + e.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(fileTxt + " Erro->" + e.Message);
            }

            //Console.WriteLine("Documento PDF criado com sucesso.");
        }

        private void AddCell(ref PdfPTable table, ref PdfPCell cell, int border, int colspan,int verticalAlignment,int horizontalAlignment, BaseColor color)
        {
            cell.Border = border;
            if (colspan > 0) cell.Colspan = colspan;
            cell.VerticalAlignment = verticalAlignment;
            cell.HorizontalAlignment = horizontalAlignment;
            cell.BackgroundColor = color;
            table.AddCell(cell);
        }

        private void PrintDocumment(string OndeGerar, string ColetaTXT, int CodVendedor, string ColetaPDFAenviar)
        {
            string enviaMensagem = "";
            int stepsEmail = 0;
            string Motivo = "";

            Seller seller = new Seller(CodVendedor);
            seller.GetSeller();

            try
            {
                Config.GetConfig();

                string sourceFile = Config.SourcePath + @"\" + ColetaPDFAenviar;
                NumOrder = NumOrder.Replace(".", "").Replace(":", "");

                string targetFile = @"F:\Aenviar\" + CnpjCpf.Replace(".", "").Replace("-", "").Replace("/", "").Replace(" ", "") +
                    @"\" + NumOrder + "-" + DateTime.Now.Hour.ToString("00") + "-" +
                    DateTime.Now.Minute.ToString("00") + "-" + DateTime.Now.Second.ToString("00") + ".pdf";

                string existeCaminhoDestino = @"F:\Aenviar\" + CnpjCpf.Replace(".", "").Replace("-", "").Replace("/", "").Replace(" ", "");
                if (!System.IO.Directory.Exists(existeCaminhoDestino))
                {
                    System.IO.Directory.CreateDirectory(existeCaminhoDestino);
                }

                //============== imprime na impressora 1 (Recepção) ======================================================================================
                if (OndeGerar == "C" || OndeGerar == "Q" || OndeGerar == "Y")
                {
                    if (File.Exists(targetFile)) { File.Delete(targetFile); }
                    try
                    {
                        File.Move(sourceFile, targetFile);
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
                    proc.StartInfo.FileName = targetFile;
                    proc.Start();
                    proc.WaitForInputIdle();
                    proc.CloseMainWindow();
                    proc.Close();
                    Console.WriteLine();
                    Console.WriteLine("Enviado o Arquivo {0} para impressora.", ColetaPDFAenviar);
                }
                else
                {
                    //================== move para o aenviar a coleta do vendedor =================================================================================
                    //System.Threading.Thread.Sleep(2000);
                    if (System.IO.File.Exists(targetFile)) { System.IO.File.Delete(targetFile); }
                    try
                    {
                        System.IO.File.Move(sourceFile, targetFile);
                        Console.WriteLine();
                        Console.WriteLine("Movido o PDF para a pasta do Cliente {0} ", CnpjCpf);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(sourceFile + " ++++> " + e.Message);
                        return;
                    }


                    if (IsNumeric(Customer.Trim()))
                    {
                        Customer = EmailCustomer;
                    }

                    
                    
                    if (IsValidEmail(EmailCustomer))
                    {

                        if (Buyer.Trim() == "") { Buyer = "Comprador"; }
                        if (EmailCustomer.Trim() == "") { EmailCustomer = "vendas@myportal.com.br"; }
                        if (Note1.Trim() != "")
                        {
                            Note1 = "<b>Ps:</b> " + Note1.Trim() + "<br />" + Note2.Trim() + "<br /><br />";
                        }
                        string Subject = NumOrder + " " + Customer;
                        enviaMensagem = EmailBody + "<br /><br />" +
                                Note1 +
                                "Atenciosamente," + "<br />" +
                                  seller.Name + "<br />" +
                                "Departamento de Vendas." + "<br />" +
                                seller.Email + "<br />" +
                                "Tel.: " +
                                seller.Phone +
                                " (31) 2101.6000  / Fax: (31) 2101.6010<br />" +
                                "Av. Bias Fortes, 1853 | B. Barro Preto | Belo Horizonte - MG | Cep 30170-012 ";

                        // envia a mensagem para o cliente
                        SendMail sendMail = new SendMail(seller.Email, enviaMensagem, EmailCustomer, Subject, targetFile, seller.Email, seller.Password, Priorities.Normal);
                        stepsEmail = sendMail.Mailing() ? 1 : 0;

                        // envia uma copia para a conta copiadeemail@myportal.com.br
                        sendMail = new SendMail(
                            seller.Email,
                            enviaMensagem,
                            emailCustomer: "copiadeemail@myportal.com.br",
                            Subject, 
                            targetFile,
                            seller.Name,
                            seller.Password,
                            Priorities.Normal
                        );
                        stepsEmail = sendMail.Mailing() ? 2 : 0;

                        // envia confirmacao para o vendedor de mensagem enviada OK
                        Subject = "OK - " + NumOrder + " " + Customer + " " + DateTime.Now.ToString();
                        enviaMensagem = NumOrder + "<br />" + CnpjCpf + "-" + Customer
                            + "<br /><b>Enviada com sucesso para:</b><br />" + EmailCustomer
                             + "<br />Em: " + DateTime.Now.ToString() + ".<br />" +
                             Note1 + ".";

                        sendMail = new SendMail(
                            emailAccount: Config.EmailAccount,
                            emailBody: enviaMensagem,
                            emailCustomer: seller.Email,
                            Subject,
                            targetFile,
                            sellerName: "Vendas",
                            password: Config.Password,
                            Priorities.Normal
                        );
                        stepsEmail = sendMail.Mailing() ? 3 : 0;

                        Console.WriteLine();
                        Console.WriteLine(Subject);
                    }
                    else
                    {
                        //caso o email do cliente nao seja valido envia um email para o vendedor avisando
                        string Subject = "((ERRO)) - " + NumOrder + " " + Customer + " " + DateTime.Now.ToString();
                        enviaMensagem = NumOrder + "<br />" + CnpjCpf + "-" + Customer + " - Vendedor: " + SalesPerson
                        + "<br /><b>NAO enviada para:</b><br />" + EmailCustomer + "<br />Motivo: <b>Email Invalido.</b>" +
                        "<br />Em: " + DateTime.Now.ToString() + ".";
                        SendMail sendMail = new SendMail(
                            emailAccount: Config.EmailAccount,
                            emailBody: enviaMensagem,
                            emailCustomer: seller.Email,
                            Subject,
                            targetFile,
                            sellerName: "Vendas",
                            password: Config.Password,
                            Priorities.High
                        );
                        stepsEmail = sendMail.Mailing() ? 3 : 0;
                    }
                }
            }
            catch (Exception e)
            {
                string Subject = "((ERRO)) - " + NumOrder + " " + Customer + " " + DateTime.Now.ToString();
                enviaMensagem = NumOrder + "<br />" + CnpjCpf + "-" + Customer + " - Vendedor: " + SalesPerson
                        + "<br /><b>NAO enviada para:</b><br />" + EmailCustomer + "<br />Motivo: " + e.Message +
                        "<br />Em: " + DateTime.Now.ToString() + ".";

                SendMail sendMail = new SendMail(
                            emailAccount: Config.EmailAccount,
                            emailBody: enviaMensagem,
                            emailCustomer: seller.Email,
                            Subject,
                            "",
                            sellerName: "Vendas",
                            password: Config.Password,
                            Priorities.High
                        );
                stepsEmail = sendMail.Mailing() ? 4 : 0;

                Motivo = e.Message;
            }
            finally
            {
                // ========= copia o arquivo txt para a pasta impressos
                string sourceFile = Path.Combine(Config.SourcePath, ColetaTXT);
                string destFile = Path.Combine(Config.TargetPath, ColetaTXT);
                try
                {
                    if (!Directory.Exists(Config.TargetPath))
                    {
                        Directory.CreateDirectory(Config.TargetPath);
                    }
                    if (File.Exists(destFile))
                    {
                        File.Delete(destFile);
                    }
                    File.Move(sourceFile, destFile);
                    Console.WriteLine("Movido arquivo de {0} para {1}", sourceFile, destFile);
                }
                catch (Exception e)
                {
                    Console.WriteLine(sourceFile + " >>> " + e.Message);
                }

                if (stepsEmail != 0)
                {
                    //gera txt com os dados do envio.
                    StreamWriter s = File.AppendText(Config.TargetPath + @"email.txt");
                    EmailCustomer = FillFields(EmailCustomer, 50);
                    CnpjCpf = (CnpjCpf.Replace(".", "").Replace("-", "").Replace("/", "").Replace(" ", "") + "00").Trim();
                    if (CnpjCpf.Length == 13)
                    {
                        CnpjCpf = CnpjCpf + "000";
                    }

                    string linha = "|" + CodVendedor.ToString("0000") + CnpjCpf + NumOrder.Substring(11, 4) + EmailCustomer +
                            DateTime.Now.ToString("ddMMyy") + stepsEmail + DateTime.Now.ToString("HHmmss") + "00";

                    if (stepsEmail == 1)
                        linha = linha + ("Enviada com sucesso para: " + EmailCustomer.Trim() + " " +
                            CnpjCpf + "-" + Customer.Trim() + " Em: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") +
                            ".=================================================").Substring(0, 100);

                    else
                        linha = linha + ("Motivo: " + Motivo + " " +
                        CnpjCpf + "-" + Customer.Trim() + " Em: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") +
                        ".==========================================================================================================" +
                        "=================================================================================================").Substring(0, 100);
                    s.WriteLine(linha);
                    s.Close();
                }
            }
        }

        private bool IsNumeric(string data)
        {
            bool isnumeric = false;
            char[] datachars = data.ToCharArray();

            foreach (var datachar in datachars)
                isnumeric = isnumeric ? char.IsDigit(datachar) : isnumeric;

            return isnumeric;
        }

        private int CountPages(string caminho)
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
                            case "C":
                                ContaLinhas++;
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

        private BaseColor GetColor(bool alternar)
        {
            if (alternar)
            {
                //CorFonte = BaseColor.WHITE;
                FontColor = BaseColor.BLACK;
                return BaseColor.WHITE;
            }
            else
            {
                FontColor = BaseColor.BLACK;
                return BaseColor.LIGHT_GRAY;
            }
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                //define a expressão regulara para validar o email
                string texto_Validar = email;
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

        private string FillFields(string campo, int tamanho)
        {
            int resto = tamanho - campo.Length;
            while (resto != 0)
            {
                campo += " ";
                resto--;
            }
            return campo;
        }

    }
}
