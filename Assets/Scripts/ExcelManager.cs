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
        //Excel�ļ�·��
        string excelFilePath = Application.streamingAssetsPath + "/" + excelFileName + ".xlsx";
        //�ļ����Ƿ���ڲ����ھʹ����ļ���
        if (!Directory.Exists(Application.streamingAssetsPath))
        {
            Directory.CreateDirectory(Application.streamingAssetsPath);
        }
        //��ȡ�ļ���Ϣ
        FileInfo fileInfo = new FileInfo(excelFileName);
        if (!fileInfo.Exists)
        {
            fileInfo = new FileInfo(excelFilePath);
        }
        //�༭�ļ���Ϣ
        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            //�༭������
            ExcelWorksheet excelWorksheet = package.Workbook.Worksheets.Add(sheetName);
            //�������
            excelWorksheet.Cells["A1"].Value = "ID";
            excelWorksheet.Cells["B1"].Value = "Name";
            excelWorksheet.Cells["C1"].Value = "Author";
            //�����������
            for (int i = 0; i < books.Count; i++)
            {
                excelWorksheet.Cells[i + 2, 1].Value = books[i].id;
                excelWorksheet.Cells[i + 2, 2].Value = books[i].name;
                excelWorksheet.Cells[i + 2, 3].Value = books[i].author;
            }
            //�洢
            package.Save();
        }
    }
}
[System.Serializable]
public class Book {
    public int id;
    public string name, author;
}