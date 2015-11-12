using System;
using System.Xml.Serialization;

namespace Учебное
{
    public class sst
    {
        [XmlElement(ElementName = "si")]
        public SiData[] si { get; set; }
    }

    public class SiData
    {
        public string t { get; set; }
    }

    [XmlRoot(Namespace = "")]
    public class worksheet
    {
        [XmlArray("sheetData")]
        [XmlArrayItem("row")]
        public Row[] Rows;
    }

    public class Row
    {
        [XmlElement("c")]
        public Cell[] FilledCells;
    }
    public class Cell
    {
        [XmlElement("v")]
        public string Value;      
    }

    public class EOData
    {
        public string ID { get; set; }
        public string Long_ID { get; set; }
        public string Name { get; set; }
        public string INN { get; set; }
        public string Agent { get; set; }
        public string Accountent { get; set; }
        public string Email { get; set; }
        public string Mobile { get; set; }
        public string ES { get; set; }
        public string PFR { get; set; }
        public string Notification { get; set; }
        public string Partner { get; set; }
        public string LastUpdate { get; set; }
        public string DateStart { get; set; }
        public string DateEnd { get; set; }
        public string Type { get; set; }
        public string Print()
        {
            return $"ID: {this.ID}{Environment.NewLine}" +
                   $"Наименование: {this.Name}{Environment.NewLine}" +
                   $"ИНН: {this.INN}{Environment.NewLine}" +
                   $"Представитель: {this.Agent}{Environment.NewLine}" +
                   $"ФИО Бухгалтера: {this.Accountent}{Environment.NewLine}" +
                   $"Email: {this.Email}{Environment.NewLine}" +
                   $"Телефон: {this.Mobile}{Environment.NewLine}" +
                   $"ЭП: {this.ES}{Environment.NewLine}" +
                   $"ПФР: {this.PFR}{Environment.NewLine}" +
                   $"Нотификация: {this.Notification}{Environment.NewLine}" +
                   $"Партнер: {this.Partner}{Environment.NewLine}" +
                   $"Последнее обновление: {this.LastUpdate}{Environment.NewLine}" +
                   $"С: {this.DateStart}{Environment.NewLine}" +
                   $"По: {this.DateEnd}{Environment.NewLine}" +
                   $"Тип: {this.Type}{Environment.NewLine}";
        }
    }
}
