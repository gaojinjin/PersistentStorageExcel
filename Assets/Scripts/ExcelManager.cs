using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System.IO;
using OfficeOpenXml;

public class ExcelManager : MonoBehaviour
{
    public List<Book> books = new List<Book>();
    public string excelFileName, sheetName;
    void Start()
    {
        //Excel文件路径
        string excelFilePath = Application.streamingAssetsPath + "/" + excelFileName + ".xlsx";
        //文件夹是否存在不存在就创建文件夹
        if (!Directory.Exists(Application.streamingAssetsPath))
        {
            Directory.CreateDirectory(Application.streamingAssetsPath);
        }
        //获取文件信息
        FileInfo fileInfo = new FileInfo(excelFileName);
        if (!fileInfo.Exists)
        {
            fileInfo = new FileInfo(excelFilePath);
        }
        //编辑文件信息
        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            //编辑工作薄
            ExcelWorksheet excelWorksheet = package.Workbook.Worksheets.Add(sheetName);
            //处理标题
            excelWorksheet.Cells["A1"].Value = "ID";
            excelWorksheet.Cells["B1"].Value = "Name";
            excelWorksheet.Cells["C1"].Value = "Author";
            //处理具体内容
            for (int i = 0; i < books.Count; i++)
            {
                excelWorksheet.Cells[i + 2, 1].Value = books[i].id;
                excelWorksheet.Cells[i + 2, 2].Value = books[i].name;
                excelWorksheet.Cells[i + 2, 3].Value = books[i].author;
            }
            //存储
            package.Save();
        }
    }
}
[System.Serializable]
public class Book {
    public int id;
    public string name, author;
}