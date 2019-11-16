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
        public bool CreateDocument(string docPath)
        {
            try
            { 
                DocumentCore document = new DocumentCore();
                document.Save(docPath);
                //System.Diagnostics.Process.Start(docPath);
                return true;
            }
            catch
            {
                return false;
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
        /// Добавлет в документ данные пользователя и дату
        /// </summary>
        /// <param name="document"></param>
        /// <param name="user"></param>
        /// <returns></returns>
        public DocumentCore ChangeDocument (DocumentCore document, User user)
        {
            Regex userRegex = new Regex(@"{user}", RegexOptions.IgnoreCase);
            Regex dateRegex = new Regex(@"{date}", RegexOptions.IgnoreCase);
            Regex tableRegex = new Regex(@"{table}", RegexOptions.IgnoreCase);
            foreach (ContentRange item in document.Content.Find(userRegex).Reverse())
            {
                item.Replace(user.FirstName + " " + user.LastName);
            }

            foreach (ContentRange item in document.Content.Find(dateRegex).Reverse())
            {
                item.Replace(DateTime.Today.ToString());
            }


            //пока костыль дл таблицы
            TableCell NewCell(int rowIndex, int colIndex)
            {
                TableCell cell = new TableCell(document);

                cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1);

                if (colIndex % 2 == 1 && rowIndex % 2 == 0 || colIndex % 2 == 0 && rowIndex % 2 == 1)
                {
                    cell.CellFormat.BackgroundColor = Color.Black;
                }

                Run run = new Run(document, string.Format("Row - {0}; Col - {1}", rowIndex, colIndex));
                run.CharacterFormat.FontColor = Color.Auto;

                cell.Blocks.Content.Replace(run.Content);

                return cell;
            }

            Table table = new Table(document, 5, 5, NewCell);

            // Place the 'Table' at the start of the 'Document'.
            // By the way, we didn't create a 'Section' in our document.
            // As we're using 'Content' property, a 'Section' will be created automatically if necessary.
            

            foreach (ContentRange item in document.Content.Find(tableRegex).Reverse())
            {
                item.Replace(table.Content);
            }

            return document;
        }        
    }
}