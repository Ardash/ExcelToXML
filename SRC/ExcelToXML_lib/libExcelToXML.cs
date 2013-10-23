using System;
using System.Data;
using System.Data.OleDb;
using System.Xml;

namespace Kovalenko.Classes {
  public class libExcelToXML {
    const string sRootEl = "root";
    const string sItemEl = "item";

    const string sId = "id";
    const string sText = "text";

    const string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                           "Data Source={0};" +
                           "Extended Properties=\"Excel 12.0; HDR=NO\"";
    const string encoding = "utf-8";

    static XmlElement findById(XmlNode e, string id) {
      XmlElement el;
      if (e.Attributes != null && 
          e.Attributes[sId] != null &&
          e.Attributes[sId].Value == id) return e as XmlElement;

      foreach (XmlNode n in e) if ((el = findById(n, id)) != null) return el;

      return null;
    }

    // Преобразование Excel-XML
    //  PathToExcelFile - имя Excel файла
    //  PathToSaveFile - имя XML файла
    //  errMsg - описание ошибки
    // RC: false - ошибка
    public static bool ExcelToXML(string PathToExcelFile, string PathToSaveFile, 
                                 ref string errMsg) {
      bool res = true;
      try {
        var s = String.Format(connStr, PathToExcelFile);
        OleDbConnection conn = new OleDbConnection(s);
        conn.Open();

        var o = new object[] { null, null, null, "TABLE" };
        DataTable sheets = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, o);
        string shName = (string)sheets.Rows[0].ItemArray[2];
        string cmdSelect = String.Format("SELECT * FROM [{0}]", shName);

        OleDbDataAdapter a = new OleDbDataAdapter(cmdSelect, conn);
        DataTable dt = new DataTable();
        a.Fill(dt);
        conn.Close();

        XmlDocument XMLForSend = new XmlDocument();

        // XML Declaration
        XmlDeclaration xmlDecl;
        xmlDecl = XMLForSend.CreateXmlDeclaration("1.0", encoding, null);
        XMLForSend.InsertBefore(xmlDecl, XMLForSend.DocumentElement);

        // Root
        XmlElement rootEl = XMLForSend.CreateElement(sRootEl);
        XMLForSend.AppendChild(rootEl);

        // Elements
        XmlElement tempRowNode;
        foreach (DataRow curRow in dt.Rows) {
          tempRowNode = XMLForSend.CreateElement(sItemEl);
          tempRowNode.SetAttribute(sId, curRow.ItemArray[0].ToString());
          tempRowNode.SetAttribute(sText, curRow.ItemArray[2].ToString());
          XmlNode e = null;
          var parent = curRow.ItemArray[1].ToString();
          if (parent != "") e = findById(rootEl, parent);
          if (e == null) e = rootEl;
          e.AppendChild(tempRowNode);
        }

        XMLForSend.Save(PathToSaveFile);
      } catch (Exception e) { res = false; errMsg = e.Message; }
      return res;
    }

  }
}

