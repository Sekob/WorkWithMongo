using MongoDB.Bson;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace UpdateDb
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

    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class Worksheet
    {
        [XmlArray("sheetData")]
        [XmlArrayItem("row")]
        public List<EOData> Rows;
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
        [XmlElement(ElementName = "c", Order = 1)]
        public string Short_Id { get; set; }
        [XmlElement(ElementName = "c", Order = 2)]
        public string _id { get; set; }
        [XmlElement(ElementName = "c", Order = 3)]
        public string Name { get; set; }
        [XmlElement(ElementName = "c", Order = 4)]
        public string INN { get; set; }
        [XmlElement(ElementName = "c", Order = 5)]
        public string Agent { get; set; }
        [XmlElement(ElementName = "c", Order = 6)]
        public string Accountent { get; set; }
        [XmlElement(ElementName = "c", Order = 7)]
        public string Email { get; set; }
        [XmlElement(ElementName = "c", Order = 8)]
        public string Mobile { get; set; }
        [XmlElement(ElementName = "c", Order = 9)]
        public string ES { get; set; }
        [XmlElement(ElementName = "c", Order = 10)]
        public string PFR { get; set; }
        [XmlElement(ElementName = "c", Order = 11)]
        public string Notification { get; set; }
        [XmlElement(ElementName = "c", Order = 12)]
        public string Partner { get; set; }
        [XmlElement(ElementName = "c", Order = 13)]
        public string LastUpdate { get; set; }
        [XmlElement(ElementName = "c", Order = 14)]
        public string DateStart { get; set; }
        [XmlElement(ElementName = "c", Order = 15)]
        public string DateEnd { get; set; }
        [XmlElement(ElementName = "c", Order = 16)]
        public string Type { get; set; }
        public string Print()
        {
            return $"ID: {this._id}{Environment.NewLine}" +
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

        public override bool Equals(object obj)
        {
            var x = obj as EOData;
            if (obj == null)
                return false;
            if (x == null)
                return false;
            if (this.Accountent != x.Accountent || this.Agent != x.Agent || this.DateEnd != x.DateEnd
                || this.DateStart != x.DateStart || this.Email != x.Email || this.ES != x.ES || this.INN != x.INN
                || this.LastUpdate != x.LastUpdate || this.Mobile != x.Mobile || this.Name != x.Name || this.Notification != x.Notification
                || this.Partner != x.Partner || this.PFR != x.PFR || this.Short_Id != x.Short_Id || this.Type != x.Type || this._id != x._id) return false;
            return true; 
        }
    }
}
