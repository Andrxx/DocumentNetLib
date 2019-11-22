using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using DocumentNetLib.BLL;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace DocumentNetLib.DAL
{
    public class DocumentManager
    {
        /// <summary>
        /// Создаем пустой документ с заданным именем в папке по уолчанию 
        /// </summary>
        /// <param name="docPath"></param>
        /// <returns></returns>
        public DocumentCore CreateDocument(string docPath)
        {
            try
            { 
                DocumentCore document = new DocumentCore();
                return document;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// Находим документ с заданным именем в папке по умолчанию, null при неудаче
        /// </summary>
        /// <param name="docPath"></param>
        /// <returns></returns>
        public DocumentCore LoadDocument(string docPath)
        {
            try
            {
                DocumentCore document = DocumentCore.Load(docPath);
                return document;
            }
            catch(Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// Сохраняет имеющийся документ с новым именем 
        /// </summary>
        /// <param name="document"></param>
        /// <param name="newName"></param>
        /// <returns></returns>
        public bool SaveDocumentAs(DocumentCore document, string newName)
        {
            try
            {
                document.Save(newName);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// Изменяет в документе текст текстовые данные в соответствии с паттерном, переданном в словаре
        /// </summary>
        /// <param name="document"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        public DocumentCore ReplaceText (DocumentCore document, Dictionary<string, string> values)
        {
            Regex tableRegex = new Regex(@"{table}", RegexOptions.IgnoreCase);
            RegexOptions options =  RegexOptions.IgnoreCase;
            foreach (KeyValuePair<string, string> pair in values)
            {
                Regex regex = new Regex(pair.Key, options);
                foreach (ContentRange item in document.Content.Find(pair.Key).Reverse())
                {
                    item.Replace(pair.Value);
                }
            }           
            return document;
        }

        /// <summary>
        /// Добавляет в документ таблицу,
        /// </summary>
        /// <param name="document">DocumentCore</param>
        /// <param name="patternName">имя паттерна для подмены в документе</param>
        /// <param name="row">количество строк</param>
        /// <param name="col">количество столбцов</param>
        /// <returns></returns>
        public DocumentCore AddTable(DocumentCore document, string patternName, int row, int col, Dictionary<int, string> dataArray)
        {
           
            TableCell NewCell(int rowIndex, int colIndex)
            {
                TableCell cell = new TableCell(document);

                cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1);
                //Run run = new Run(document, string.Format("Row - {0}; Col - {1}", rowIndex, colIndex));
                Run run = new Run(document);
                run.CharacterFormat.FontColor = Color.Auto;
                if (dataArray != null)
                {
                    try
                    {
                        int key = int.Parse(rowIndex.ToString() + colIndex.ToString());
                        run.Text = dataArray[key];
                    }
                    catch { }
                }
                
                cell.Blocks.Content.Replace(run.Content);
                return cell;
            }

            

            Table table = new Table(document, row, col, NewCell);

            foreach (ContentRange item in document.Content.Find(patternName).Reverse())
            {
                item.Replace(table.Content);
            }

            return document;
        }

        /// <summary>
        /// метод тестировки - содание шаблона
        /// </summary>
        /// <returns></returns>
        public DocumentCore createTemplate()
        {
            DocumentCore document = new DocumentCore();

            // Add new section.
            Section section = new Section(document);
            document.Sections.Add(section);

            // Let's set page size A4.
            section.PageSetup.PaperType = PaperType.A4;

            // Add two paragraphs using different ways:
            // Way 1: Add 1st paragraph.
            Paragraph par1 = new Paragraph(document);
            par1.ParagraphFormat.Alignment = HorizontalAlignment.Left;
            section.Blocks.Add(par1);

            // Let's create a characterformat for text in the 1st paragraph.
            CharacterFormat cf = new CharacterFormat() { FontName = "Times New Roman", Size = 14, FontColor = Color.Black };

            Run text1 = new Run(document, "Приветствую {user}!");
            text1.CharacterFormat = cf;
            par1.Inlines.Add(text1);

            // Let's add a line break into our paragraph.
            par1.Inlines.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));
            par1.Inlines.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));

            Paragraph par2 = new Paragraph(document);
            par2.ParagraphFormat.Alignment = HorizontalAlignment.Center;
            section.Blocks.Add(par2);
            document.Content.End.Insert(@"Просто текст.", new CharacterFormat()
            { FontName = "Times New Roman", Size = 14, FontColor = Color.Black});
            par2.Inlines.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));

            Paragraph par3 = new Paragraph(document);
            par3.ParagraphFormat.Alignment = HorizontalAlignment.Left;
            section.Blocks.Add(par3);
            document.Content.End.Insert(@"Lorem Ipsum - это текст-рыба, часто используемый в печати и вэб-дизайне. Lorem Ipsum " +
                "является стандартной рыбой для текстов на латинице с начала XVI века. В то время некий безымянный печатник создал большую коллекцию" +
                " размеров и форм шрифтов, используя Lorem Ipsum для распечатки образцов. Lorem Ipsum не только успешно пережил без заметных изменений пять веков, " +
                "но и перешагнул в электронный дизайн. Его популяризации в новое время послужили публикация листов Letraset с образцами Lorem Ipsum в 60-х годах и, " +
                "в более недавнее время, программы электронной вёрстки типа Aldus PageMaker, в шаблонах которых используется Lorem Ipsum.", 
                new CharacterFormat() { FontName = "Times New Roman", Size = 14, FontColor = Color.Black });

            Paragraph par4 = new Paragraph(document);
            par4.ParagraphFormat.Alignment = HorizontalAlignment.Left;
            section.Blocks.Add(par4);
            document.Content.End.Insert(@"{user}									{date}",
                new CharacterFormat() { FontName = "Times New Roman", Size = 14, FontColor = Color.Black });
            par4.Inlines.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));

            Paragraph par5 = new Paragraph(document);
            par5.ParagraphFormat.Alignment = HorizontalAlignment.Center;
            section.Blocks.Add(par5);
            document.Content.End.Insert(@"Здесь будет заголовок таблицы.",
                new CharacterFormat() { FontName = "Times New Roman", Size = 14, FontColor = Color.Black });
            par5.Inlines.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));

            Paragraph par6 = new Paragraph(document);
            par6.ParagraphFormat.Alignment = HorizontalAlignment.Left;
            section.Blocks.Add(par6);
            document.Content.End.Insert(@"{table}",
                new CharacterFormat() { FontName = "Times New Roman", Size = 14, FontColor = Color.Black });
            par6.Inlines.Add(new SpecialCharacter(document, SpecialCharacterType.LineBreak));

            return document;
        }
    }
}