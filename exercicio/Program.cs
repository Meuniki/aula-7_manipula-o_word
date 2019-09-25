using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace exercicio
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Exercicio
                Document exercicioDoc = new Document ();
            #endregion

            #region SesacaoCapa
                Section secaoCapa = exercicioDoc.AddSection();
            #endregion

            #region Paragrafo01
                Paragraph titulo = secaoCapa.AddParagraph ();
            #endregion

            #region Titulo
                titulo.AppendText("Exercício Concluido\n\n");
            #endregion
        }
    }
}
