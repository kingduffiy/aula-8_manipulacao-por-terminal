using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace aula_8_manipulacao_por_terminal {
    class Program {
        static void Main (string[] args) {
            #region Criacao do documento
            // Cria um documento com o nome exemploDoc
            Document exemploDoc = new Document ();
            #endregion

            #region Criacao de secao no documento
            // Adiciona uma seção com o nome secaoCapa ao documento
            // Cada seção pode ser entendida como uma pagina do documento
            Section secaocapa = exemploDoc.AddSection ();
            #endregion

            #region Criar um paragrafo
            // Cria um paragrafo com o nome e titulo e adiciona a seção secaoCapa
            // Os paragrafos são necessários para inserção de textos, imagens, tabelas etc
            Paragraph titulo = secaocapa.AddParagraph ();
            #endregion

            #region Adiciona texto ao paragrafo
            // Adiciona o texto Exemplo de titulo ao paragrafo titulo
            titulo.AppendText ("Exemplo de título\n\n");
            #endregion

            #region Formatar paragrafo
            // Através da propriedade HorizontalAlignment, é possivel alinhar o parágrafo
            titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

            //Cria um estilo com o nome estilo01 e adiciona ao documento
            ParagraphStyle estilo01 = new ParagraphStyle (exemploDoc);

            //Adiciona um nome ao estilo01
            estilo01.Name = "Cor do Titulo";

            //Definir a cor do titulo
            estilo01.CharacterFormat.TextColor = Color.DarkBlue;

            //Define que o texto será em negrito
            estilo01.CharacterFormat.Bold = true;

            //Adiciona o estilo01 ao documento exemploDoc
            exemploDoc.Styles.Add (estilo01);

            // Aplica o estilo01 ao paragrafo do titulo
            titulo.ApplyStyle (estilo01.Name);

            #region Trabalhahr com tabulacao
            // Adiciona um paragrafo textoCapa a seção secaocapa
            Paragraph textoCapa = secaocapa.AddParagraph ();

            //Adiciona um texto ao paragrafo com tabulação
            textoCapa.AppendText ("\tEste é um exemplo de texto com tabulação\n");

            //Adiciona um novo paragrafo a mesma seção (secaocapa)
            Paragraph textoCapa2 = secaocapa.AddParagraph ();

            textoCapa2.AppendText ("\tBasicamente, então, uma seção representa uma página do documento e os paragrafos dentro de uma mesma seção," + "Obviamente, aparecem na mesma seção");
            #endregion

            #region Inserir imagens

            //Adiciona um paragráfo a seção capa
            Paragraph imagemCapa = secaocapa.AddParagraph ();

            //Adiciona um texto ao paragrafo imagemCapa
            imagemCapa.AppendText ("\n\n\tAgora vamos inserir uma imagem ao documento\n\n");

            //Centraliza
            imagemCapa.Format.HorizontalAlignment = HorizontalAlignment.Center;

            //Adiciona uma imagem com o nome imagemExemplo ao parágrafo imagemCapa
            DocPicture imagemExemplo = imagemCapa.AppendPicture (Image.FromFile (@"saida\img\logo_csharp.png"));

            //Define uma largura e uma altura para a imagem
            imagemExemplo.Width = 300;
            imagemExemplo.Height = 300;

            #endregion

            #region Adicionar nova seção
            //Adiciona uma nova seção
            Section secaoCorpo = exemploDoc.AddSection ();

            // Adiciona um paragrafo à seção secaoCorpo
            Paragraph paragrafoCorpo1 = secaoCorpo.AddParagraph ();

            //Adiciona um texto ao parágrafo paragrafoCorpo1
            paragrafoCorpo1.AppendText ("\tEste é um exemplo de parágrafo criado em uma nova seção." + "\tComo foi criada uma nova seção, perceba que esse texto aparece em uma nova pagina");

            #endregion

            #region Adicionar uma tabela
                //Adiciona uma tabela a seção secaoCorpo
                Table tabela = secaoCorpo.AddTable(true);

                //Cria o cabeçalho da tabela
                String[] cabecalho = {"Item", "Descrição", "Qtd", "Preço Unit.","Preço"};

                String[][] dados = {
                    new String[]{"Cenoura","Vegetal muito nutritivo","1","R$4,00", "R$4,00"},
                    new String[]{"Batata","Vegetal muito consumido","2","R$5,00", "R$10,00"},
                    new String[]{"Alface","Vegetal utilizado desde 500 a.C.","2","R$2,00", "R$4,00"},
                    new String[]{"Sopa de Macaco","Uma delicia","1","R$30,00", "R$30,00"},
                };
                
                //Adiciona as células na tabela
                tabela.ResetCells(dados.Length + 1, cabecalho.Length);
                
                //Adiciona uma  linha na posição [0] do vetor de linhas e define que esta linha é o cabeçalho
                TableRow Linha1 = tabela.Rows[0];
                Linha1.IsHeader = true;

                //Define a altura da linha
                Linha1.Height = 23;

                //Formatação do cabeçalho
                Linha1.RowFormat.BackColor = Color.AliceBlue;

                for (int i = 0; i < cabecalho.Length; i++)
                {   
                    //Alinhamento das células
                    Paragraph p = Linha1.Cells[i].AddParagraph();
                    Linha1.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    
                    // Formatação dos dados do cabeçalho
                    TextRange TR = p.AppendText(cabecalho[i]);
                    TR.CharacterFormat.FontName = "Calibri";
                    TR.CharacterFormat.FontSize = 14;
                    TR.CharacterFormat.TextColor = Color.Teal;
                    TR.CharacterFormat.Bold = true;
                }


                // Adiciona as linhas do corpo da tabela
                for (int r = 0; r < dados.Length; r++)
                {
                    TableRow LinhaDados = tabela.Rows[r + 1];

                    //Define a altura da linha
                    LinhaDados.Height = 20;

                    //Percorre as colunas
                    for (int c = 0; c < dados[r].Length; c++)
                    {

                        //Alinha as células
                        LinhaDados.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                        //Preenche os dados nas linhas
                        Paragraph p2 = LinhaDados.Cells[c].AddParagraph();
                        TextRange TR2 = p2.AppendText(dados[r][c]);

                        //Formata as células
                        p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR2.CharacterFormat.FontName = "Calibri";
                        TR2.CharacterFormat.FontSize = 12;
                        TR2.CharacterFormat.TextColor = Color.Brown;
                        
                    }
                }
            #endregion
            #region Salvar arquivo
            // Salva o carquivo em .Docx
            //Utiliza o metodo SaveToFile para salvar o arquivo no formato desejado
            // Assim como no Word, caso ja existia um aquivo com ese nome, é substituido
                exemploDoc.SaveToFile(@"saida\exemplo_arquivo_word.docx",FileFormat.Docx);
            #endregion

            #endregion
        }
    }
}