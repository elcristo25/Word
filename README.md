public void AñadirEncabezado(string nombreDocumento, string contenidoEncabezado)
{
     this.TieneEncabezado(nombreDocumento);
     using (WordprocessingDocument documento = WordprocessingDocument.Open(nombreDocumento, true))
     {

           if (this.tieneEncabezado)
           {
               return;
           }
           try { documento.MainDocumentPart.DeletePart("HeaderId1"); }
           catch (ArgumentOutOfRangeException) { }
     
           HeaderPart ParteEncabezado = documento.MainDocumentPart.AddNewPart<HeaderPart>("HeaderId1");
     
           Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
           header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
           header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
           header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
           header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
           header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
           header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
           header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
           header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
           header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
           header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
           header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
           header1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
           header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
           header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
           header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
           header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
     
           Paragraph paragraph3 = new Paragraph() { };
     
           ParagraphProperties paragraphProperties3 = new ParagraphProperties();
           ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "Encabezado" };
     
           paragraphProperties3.Append(paragraphStyleId3);
     
           Run run2 = new Run();
           Text text2 = new Text();
           text2.Text = contenidoEncabezado;
     
           run2.Append(text2);
     
           paragraph3.Append(paragraphProperties3);
           paragraph3.Append(run2);
     
           Paragraph paragraph4 = new Paragraph() { };
     
           ParagraphProperties paragraphProperties4 = new ParagraphProperties();
           ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "Encabezado" };
     
           paragraphProperties4.Append(paragraphStyleId4);
     
           paragraph4.Append(paragraphProperties4);
     
           header1.Append(paragraph3);
           header1.Append(paragraph4);
     
           ParteEncabezado.Header = header1;
     
           documento.MainDocumentPart.AddPart(ParteEncabezado);
     
           HeaderReference headerReference1 = new HeaderReference() { Type = HeaderFooterValues.Default, Id = "HeaderId1" };
           int contadorHijosBody = 0;
           foreach (var prueba in documento.MainDocumentPart.Document.Body.ChildElements)
           {
               if (prueba.LocalName == "sectPr")
               {
                   documento.MainDocumentPart.Document.Body.ChildElements[contadorHijosBody].ReplaceChild(headerReference1, documento.MainDocumentPart.Document.Body.ChildElements[contadorHijosBody].ChildElements[0]);
               }
               contadorHijosBody++;
           }


     }
}
