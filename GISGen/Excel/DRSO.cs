using System;
using System.Collections.Generic;

namespace GISGen.Excel
{
    public class DRSO
    {
        public Excel.Document Document;

        public DRSO(string fileName)
        {
            Document = new Excel.Document(fileName);
            var conf = Document.Worksheets[12];
            if (conf.Cells["B1"].Text != "DRSO")
                throw new DRSOException("Неверный файл шаблона");
            if (conf.Cells["B2"].Text != "9.0.1.2")
                throw new DRSOException("Данная версия шаблона не поддерживается");
        }

        public void Close()
        {
            Document.Close();
        }

        public Dictionary<int, string> DRSOWorksheets = new Dictionary<int, string>()
        {
            {1,"Договоры ресурсоснабжения"},
            {2,"Предметы договоров"},
            {3,"Объекты жилищного фонда"},
            {4,"КУ и КР по ОЖФ"},
            {5,"Показатели качества КР"},
            {6,"Иные показатели качества"},
            {7,"Температурный график"}
        };

    }

    // Специальное исключение
    [Serializable]
    public class DRSOException : ApplicationException
    {
        public DRSOException() { }
        public DRSOException(string message) : base(message) { }
        public DRSOException(string message, Exception ex) : base(message) { }
        // Конструктор для обработки сериализации типа
        protected DRSOException(System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext contex)
            : base(info, contex)
        { }
    }

}