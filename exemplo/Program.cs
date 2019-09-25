using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace exemplo {
    class Program {
        static void Main (string[] args) {
            #region Criacao do documento
            // Cria um documento com o nome exemploDoc
            Document exemploDoc = new Document ();
            #endregion

            #region Cria secao no documento
            //Adiciona uma seção com o nome secaocapa ao documento
            //Cada seção pode ser entendida como uma pagina do documento
            Section secaoCapa = exemploDoc.AddSection ();
            #endregion

            #region Criar paragrafo
            //Cria um paragrafo com o nome titulo adiciona à secaoCapa
            // Os paragrafos são nescesários para incerção de texto, imagen, tabelas etc
            Paragraph titulo = secaoCapa.AddParagraph ();
            #endregion

            #region Adiciona texto ao paragrafo
            // Adiciona o texto Exemplo de título ao paragrafo titulo
            titulo.AppendText ("Exemplo de Título\n\n");
            #endregion

            #region Formatar paragrafo
            // Atraves da propriedade HorizontalAlignment,é possível alinhar o parágrafo
            titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

            // Cria um estilo com o nome estilo01 e adiciona ao documento
            ParagraphStyle estilo01 = new ParagraphStyle (exemploDoc);

            // Defini um nome aao estilo01
            estilo01.Name = "Cor do titulo";

            // Definir a cor do titulo
            estilo01.CharacterFormat.TextColor = Color.DarkBlue;

            // Define que o texto será em negrito
            estilo01.CharacterFormat.Bold = true;

            // Adiciona o estilo01  ao documento exemploDoc
            exemploDoc.Styles.Add (estilo01);

            // Aplica o estilo01 ao parágrafo titulo
            titulo.ApplyStyle (estilo01.Name);
            #endregion

            #region Trabalhar com tabulação
            // Adiciona um paragrafo textoCapa á seção sacaoCapa
            Paragraph textoCapa = secaoCapa.AddParagraph ();

            // Adiciona um texto ao parágrafo com tabulação
            textoCapa.AppendText ("\tEste é um exemplo de texto com tabulação\n");

            //Adiciona um novo paragrafo a mesma seção (secaoCapa)
            Paragraph textoCapa2 = secaoCapa.AddParagraph ();

            // Adiciona um texto ao parágrafo textoCApa2 com concatenação
            textoCapa2.AppendText ("\tBasicamente, então, uma seção representa uma pagina do documento e os paragrafos dentro de uma mesma seção," + "Obiviamente, aparecen na mesma página");
            #endregion

            #region Inserir Imagens
            // Adiciona uma imagen a seção capa
            Paragraph imagenCapa = secaoCapa.AddParagraph ();

            // Adiciona um texto ao parágrafo imagemCapa
            imagenCapa.AppendText ("\n\n Agora vamos inserir uma imagen ao documento\n\n");
            imagenCapa.Format.HorizontalAlignment = HorizontalAlignment.Center;

            // Adiciona uma imagem com o nome imagemExemplo ao parágrafo imagrmCapa
            DocPicture imagemExeoplo = imagenCapa.AppendPicture (Image.FromFile (@"saida\img\logo_csharp.png"));

            //Define uma largura e uma altura para a imagem
            imagemExeoplo.Width = 300;
            imagemExeoplo.Height = 300;

            #endregion

            #region Adiciona nova secao

            // Adiciona uma nova seção
            Section secaoCorpo = exemploDoc.AddSection ();

            // Adiciona um parágrafo à seção secaocorpo
            Paragraph paragrafoCorpo1 = secaoCorpo.AddParagraph ();

            paragrafoCorpo1.AppendText ("\tEsse é uma paragrafo criado em uma nova seção" + "\tComo foi criada uma nova seção perceba que esse texto aparece em uma nova página");
            #endregion

            #region Adicionar tabela
            //Adiciona uma tabela a seção secaCorpo
            Table tabela = secaoCorpo.AddTable (true);

            // Cria o cabeçalho da tabela
            String[] cabecalho = { "item", "Descrição", "Qtd.", "Preço Unit.", "Preço" };

            String[][] dados = {
                new String[] { "Cenoura", "Vegetal muito nutritivo", "1", "R$4,00", "R$4,00" },
                new String[] { "Batata", "Vegetal muito consumido", "2", "R$5,00", "R$10,00" },
                new String[] { "Alface", "Vegetal ultilizado desde 500 a.C.", "1", "R$1,50", "R$1,50" },
                new String[] { "Tomate", "Tomate é uma fruta", "2", "R$6,00", "R$12,00" },
            };

            // Adiciona as células na tabela
            tabela.ResetCells (dados.Length + 1, cabecalho.Length);

            // Adiciona uma lonha na posição 0 no vetor de linhas
            // E define que esta linha é o caceçalho
            TableRow Linha1 = tabela.Rows[0];
            Linha1.IsHeader = true;

            // DEfine a altura das linhas
            Linha1.Height = 23;

            // Formatação do cabeçalho
            Linha1.RowFormat.BackColor = Color.AliceBlue;

            // Percorre as colunas do cabeçalho
            for (int i = 0; i < cabecalho.Length; i++) {
                // Alinhamento das células
                Paragraph p = Linha1.Cells[i].AddParagraph ();
                Linha1.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                p.Format.HorizontalAlignment = HorizontalAlignment.Center;

                //Formatar os dado do cabeçalho
                TextRange TR = p.AppendText (cabecalho[i]);
                TR.CharacterFormat.FontName = "Calibri";
                TR.CharacterFormat.FontSize = 14;
                TR.CharacterFormat.TextColor = Color.Teal;
                TR.CharacterFormat.Bold = true;
            }

            //Adiciona as linhas do corpo da tabela
            for (int r = 0; r < dados.Length; r++) {
                TableRow LinhaDados = tabela.Rows[r + 1];

                // Define a altura da linha
                LinhaDados.Height = 20;

                // Percorre as colunas
                for (int c = 0; c < dados[r].Length; c++) {
                    LinhaDados.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                    //Preencher os dados nas linhas
                    Paragraph p2 = LinhaDados.Cells[c].AddParagraph ();
                    TextRange TR2 = p2.AppendText (dados[r][c]);

                    //Formatar as células
                    p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    TR2.CharacterFormat.FontName = "Calibri";
                    TR2.CharacterFormat.FontSize = 12;
                    TR2.CharacterFormat.TextColor = Color.Brown;
                    TR2.CharacterFormat.Bold = true;

                }
            }
            #endregion

            #region Salvar arquivo
            // Salva arquivo em .docx
            // Ultiliza o metodo SAveToFile para salvar o arquivo no formato desejado
            // Assim como no Word, caso já exixta um arquivo com este, nome, é substituído
            exemploDoc.SaveToFile (@"saida\exemplo_arquivo_word.docx", FileFormat.Docx);
            #endregion

        }
    }
}