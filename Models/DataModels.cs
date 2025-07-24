using System;
using System.Collections.Generic;

namespace TaskBoard.Models
{
    public class TabData
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Name { get; set; } = "新标签页";
        public string DotColor { get; set; } = "#757575"; // 默认灰色
        public DateTime LastModified { get; set; } = DateTime.Now;
        public List<ColumnDefinition> Columns { get; set; } = new List<ColumnDefinition>();
        public List<RowData> Rows { get; set; } = new List<RowData>();
    }

    public class ColumnDefinition
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Name { get; set; } = "新列";
        public ColumnType Type { get; set; } = ColumnType.Text;
        public List<OptionItem> Options { get; set; } = new List<OptionItem>();
    }

    public class OptionItem
    {
        public string Label { get; set; } = "";
        public string Color { get; set; } = "#FFFFFF";
    }

    public class RowData
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public Dictionary<string, object> Data { get; set; } = new Dictionary<string, object>();
    }

    public enum ColumnType
    {
        Text,
        Single,
        Multi,
        Image,
        TodoList,
        TextArea  // 文本域类型
    }

    public class TodoItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Text { get; set; } = "";
        public bool IsCompleted { get; set; } = false;
    }

    public class AppData
    {
        public List<TabData> Tabs { get; set; } = new List<TabData>();
    }
}