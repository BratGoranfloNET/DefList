using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Newtonsoft.Json;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Text;
using System.Web;
using Microsoft.Office.Server.UserProfiles;

namespace Sibur.SharePoint.DefList
{
    class DocUtil
    {

        private string TryGetProperty(UserProfile profile, string propertyName)
        {
            string result = String.Empty;

            UserProfileValueCollection property = null;

            try
            {
                property = profile[propertyName];
                result = property.Value?.ToString();
            }
            catch
            {
            }

            return result;
        }


        public static SPList GetListByRelariveUrlOrId(SPWeb web, string listRelativeUrlOrID)
        {
            SPList list = null;
            Guid listID = Guid.Empty;
            Guid.TryParse(listRelativeUrlOrID, out listID);
            if (listID != Guid.Empty)
            {
                list = web.Lists.GetList(listID, true);
            }
            else
            {
                list = web.GetList(SPUrlUtility.CombineUrl(web.ServerRelativeUrl, listRelativeUrlOrID));
            }
            return list;
        }


        private static Cell ConstructCell(string value, CellValues dataType, Nullable<UInt32> styleIndex = null)
        {
            Cell cell1 = new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };

            if (styleIndex != null)
                cell1.StyleIndex = styleIndex;

            return cell1;
        }


        private static Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;

            DocumentFormat.OpenXml.Spreadsheet.Fonts fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 11 },
                    new FontName() { Val = "Times New Roman" }


                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 12 },
                    new DocumentFormat.OpenXml.Spreadsheet.Bold()
                // new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = "FFFFFF" }

                ),
                new Font( // Index 2 -
                    new FontSize() { Val = 12 },
                     new FontName() { Val = "Times New Roman" },
                    new DocumentFormat.OpenXml.Spreadsheet.Bold()


                ),
                new Font( // Index 3 
                    new FontSize() { Val = 12 },
                     new FontName() { Val = "Times New Roman" }
                // new DocumentFormat.OpenXml.Spreadsheet.Bold()
                ),                
                new Font( // Index 4 
                    new FontSize() { Val = 14 },
                    new FontName() { Val = "Times New Roman" },
                    new DocumentFormat.OpenXml.Spreadsheet.Bold()
                )
                );

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "ff08c0ce" } })
                    { PatternType = PatternValues.Solid }), // Index 2 - header

                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "ff97ffdc" } })
                    { PatternType = PatternValues.Solid }), // Index 3 - top level process

                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "ffb9ffdc" } })
                    { PatternType = PatternValues.Solid }), // Index 4 - group level process

                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "ffdcffb9" } })
                    { PatternType = PatternValues.Solid }) // Index 5 -  process

                );

            Borders borders = new Borders(
                    new DocumentFormat.OpenXml.Spreadsheet.Border(), // index 0 default
                    new DocumentFormat.OpenXml.Spreadsheet.Border( // index 1 black border
                        new LeftBorder(new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );


            CellFormats cellFormats = new CellFormats(
                    new CellFormat(new Alignment() {  Vertical = VerticalAlignmentValues.Center, WrapText = true }), // default
                    new CellFormat { Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }, FontId = 0, FillId = 0, BorderId = 0, ApplyBorder = true, ApplyAlignment = true }, // body
                    new CellFormat { Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }, FontId = 2, FillId = 0, BorderId = 1, ApplyFill = true },   
                    new CellFormat { Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }, FontId = 3, FillId = 0, BorderId = 1, ApplyFill = true },   // 3
                    new CellFormat { Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }, FontId = 2, FillId = 0, BorderId = 1, ApplyFill = true },   // 4
                    new CellFormat { Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, WrapText = true }, FontId = 0, FillId = 0, BorderId = 0, ApplyFill = true  },   //5
                    new CellFormat { Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }, FontId = 4, FillId = 0, BorderId = 0, ApplyFill = true },   // 6
                    new CellFormat { Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }, FontId = 3, FillId = 0, BorderId = 0, ApplyFill = true },  // 7
                    new CellFormat { Alignment = new Alignment() { WrapText = true }, FontId = 0, FillId = 0, BorderId = 0, ApplyFill = true }   // process

                );


            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;

        }
        
        public void GenerateExcelFile(string listDefUrlOrID, string listDefItemId, string listOrgUrlOrID, string listOrgItemId, ResponseModel model)
        {
            try
            {
                string position1 = string.Empty; 
                string position2 = string.Empty;
                string position3 = string.Empty;

                string name1 = "_____________________"; 
                string name2 = "_____________________"; 
                string name3 = string.Empty;

                string date1 = @"_____ _______________2020 г.";
                string date2 = @"_____ _______________2020 г.";

                string defId = string.Empty;
                string defTitle = string.Empty;
                string lookUpValue = string.Empty;
                bool isLookUpNotNull = false;

                var siteId = SPContext.Current.Site.ID;
                var webId = SPContext.Current.Web.ID;

                model.SiteId = siteId;
                model.WebId = webId;

                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    using (SPSite site = new SPSite(siteId))
                    using (SPWeb web = site.OpenWeb(webId))
                    {
                        model.SiteUrl = site.Url;
                        model.WebUrl = web.Url;

                        Divisions division = new Divisions();

                        SPList defList = GetListByRelariveUrlOrId(web, listDefUrlOrID);

                        SPServiceContext serviceContext = SPServiceContext.GetContext(site);
                        UserProfileManager manager = new UserProfileManager(serviceContext);

                        if (defList != null)
                        {
                          model.DefListUrl = defList. DefaultViewUrl;
                          SPListItem defListItem = defList.GetItemById(Convert.ToInt32(listDefItemId));
                            if (defListItem != null)
                            {
                                SPFieldUserValue userEditor = new SPFieldUserValue(web, defListItem["Editor"].ToString());
                                SPUser spUser3 = userEditor.User;
                                if (spUser3 != null)
                                {
                                    string[] temp = spUser3.Name.Split(' ');
                                    name3 += temp[0] + " " + temp[1].Substring(0, 1) + " " + temp[2].Substring(0, 1);

                                    UserProfile profile3 = manager.GetUserProfile(spUser3.LoginName);
                                    position3 = TryGetProperty(profile3, "title");
                                }
                                
                                defId = defListItem.ID.ToString(); 
                                defTitle =  defListItem["Title"].ToString();

                                model.ItemUrl = defListItem.Url;
                                string jsonValue = defListItem["ListDataJSON"].ToString();
                                model.JsonValue = jsonValue;
                                division = JsonConvert.DeserializeObject<Divisions>(jsonValue);
                                                                
                                var templookUp = defListItem["Условие стеснённости"];
                                if(templookUp != null)
                                {
                                    SPFieldLookupValue lookupField = new SPFieldLookupValue(templookUp.ToString());
                                    if (lookupField != null)
                                    {
                                        lookUpValue = lookupField.LookupValue;
                                        isLookUpNotNull = true;
                                    }
                                    
                                }
                            }    
                        }                                               

                        
                        SPList orgList = GetListByRelariveUrlOrId(web, listOrgUrlOrID);
                        if (orgList != null)
                        {
                            SPListItem orgListItem = orgList.GetItemById(Convert.ToInt32(listOrgItemId));

                            SPFieldUserValue userValue1 = new SPFieldUserValue(web, orgListItem["fldApprove"].ToString());
                            SPUser spUser1 = userValue1.User;
                            if (spUser1 != null)
                            {
                                string[] temp = spUser1.Name.Split(' ');
                                name1 += temp[0] + " " + temp[1].Substring(0,1) + " " + temp[2].Substring(0,1);

                                UserProfile profile1 = manager.GetUserProfile(spUser1.LoginName);
                                position1 = TryGetProperty(profile1, "title");
                            }

                            SPFieldUserValue userValue2 = new SPFieldUserValue(web, orgListItem["fldConfirm"].ToString());
                            SPUser spUser2 = userValue2.User;
                            if (spUser2 != null)
                            {
                                string[] temp = spUser2.Name.Split(' ');
                                name2 += temp[0] + " " + temp[1].Substring(0, 1) + " " + temp[2].Substring(0, 1);

                                UserProfile profile2 = manager.GetUserProfile(spUser2.LoginName);
                                position2 = TryGetProperty(profile2, "title");
                            }

                        }


                        string tempFileUrl = SPUtility.GetCurrentGenericSetupPath(@"TEMPLATE\LAYOUTS\Sibur.SharePoint.DefList\Template\DefectiveListTemplate.xlsx");                        
                        model.FileExist = File.Exists(tempFileUrl);

                        // SPFile templateFile = web.GetFile(tempFileUrl);
                        // byte[] byteArray = templateFile.OpenBinary();
                        byte[] byteArray = File.ReadAllBytes(tempFileUrl);

                        model.ArrayLength = byteArray.Length;
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            memoryStream.Write(byteArray, 0, byteArray.Length);
                            using (SpreadsheetDocument spreadDocument = SpreadsheetDocument.Open(memoryStream, true))
                            {
                                WorkbookPart workBookPart = spreadDocument.WorkbookPart;
                                Sheet sheet = spreadDocument.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                                Worksheet worksheet = (spreadDocument.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                                
                                if (workBookPart.WorkbookStylesPart.Stylesheet.Descendants<StylesheetExtensionList>().Any())
                                {
                                    workBookPart.WorkbookStylesPart.Stylesheet.RemoveAllChildren<StylesheetExtensionList>();
                                    workBookPart.WorkbookStylesPart.Stylesheet.RemoveAllChildren<Stylesheet>();
                                }

                                WorkbookStylesPart workbookStylesPart = workBookPart.WorkbookStylesPart; 

                                Stylesheet stylesheet1 = GenerateStylesheet();
                                workbookStylesPart.Stylesheet = stylesheet1;
                                                               

                                Row row0 = new Row();
                                row0.Append(
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("СОГЛАСОВАНО", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("УТВЕРЖДАЮ", CellValues.String, 5),
                                   ConstructCell("", CellValues.String, 0));


                                Row row1 = new Row();
                                row1.Append(
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell(position1, CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell(position2, CellValues.String,5),
                                   ConstructCell("", CellValues.String, 0));

                                Row row2 = new Row();
                                row2.Append(
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell(name1, CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell(name2, CellValues.String, 5),
                                   ConstructCell("", CellValues.String, 0));

                                Row row3 = new Row();
                                row3.Append(
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell(date1, CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell(date2, CellValues.String, 5),
                                   ConstructCell("", CellValues.String, 0));


                                Row row35 = new Row();
                                row35.Append(
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1));



                                Row row4 = new Row();
                                row4.Append(
                                   ConstructCell("", CellValues.String, 5),
                                   ConstructCell("Ведомость объемов работ № " + defTitle, CellValues.String, 6),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1));

                                Row row5 = new Row();
                                row5.Append(
                                   ConstructCell("", CellValues.String, 5),
                                   ConstructCell("", CellValues.String, 6),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1),
                                   ConstructCell("", CellValues.String, 1));

                                Row row6 = new Row();
                                row6.Append(
                                   ConstructCell("", CellValues.String, 2),
                                   ConstructCell("п/п", CellValues.String, 2),
                                   ConstructCell("Наименование выполняемой работы", CellValues.String, 2),
                                   ConstructCell("Ед.изм.", CellValues.String, 2),
                                   ConstructCell("Кол-во", CellValues.String, 2),
                                   ConstructCell("МТР", CellValues.String, 2),
                                   ConstructCell("Примечание", CellValues.String, 2));

                                Row row7 = new Row();
                                row7.Append(
                                   ConstructCell("", CellValues.String, 2),
                                   ConstructCell("1", CellValues.Number, 2),
                                   ConstructCell("2", CellValues.Number, 2),
                                   ConstructCell("3", CellValues.Number, 2),
                                   ConstructCell("4", CellValues.Number, 2),
                                   ConstructCell("5", CellValues.Number, 2),
                                   ConstructCell("6", CellValues.Number, 2));

                                sheetData.AppendChild(row0);
                                sheetData.AppendChild(row1);
                                sheetData.AppendChild(row2);
                                sheetData.AppendChild(row3);
                                sheetData.AppendChild(row35);
                                sheetData.AppendChild(row4);
                                sheetData.AppendChild(row5);
                                sheetData.AppendChild(row6);
                                sheetData.AppendChild(row7);


                                ////////////////////////MERGE HEAD////////////////////////
                                MergeCells mergeCells;

                                if (worksheet.Elements<MergeCells>().Count() > 0)
                                    mergeCells = worksheet.Elements<MergeCells>().First();
                                else
                                {
                                    mergeCells = new MergeCells();
                                  
                                    if (worksheet.Elements<CustomSheetView>().Count() > 0)
                                    {
                                        worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                                    }
                                    else
                                    {
                                        var sss = worksheet.Elements<SheetData>().Count();
                                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                                    }
                                }
                                
                                MergeCell mergeCell1 = new MergeCell()
                                {
                                    Reference =
                                    new StringValue("B2:C2")
                                };

                                MergeCell mergeCell2 = new MergeCell()
                                {
                                    Reference =
                                    new StringValue("B3:C3")
                                };

                                MergeCell mergeCell3 = new MergeCell()
                                {
                                    Reference =
                                    new StringValue("F2:G2")
                                };

                                MergeCell mergeCell4 = new MergeCell()
                                {
                                    Reference =
                                    new StringValue("F3:G3")
                                };

                                //
                                MergeCell mergeCell5 = new MergeCell()
                                {
                                    Reference =
                                    new StringValue("B4:C4")
                                };

                                MergeCell mergeCell6 = new MergeCell()
                                {
                                    Reference =
                                    new StringValue("B5:C5")
                                };

                                MergeCell mergeCell7 = new MergeCell()
                                {
                                    Reference =
                                    new StringValue("F4:G4")
                                };

                                MergeCell mergeCell8 = new MergeCell()
                                {
                                    Reference =
                                    new StringValue("F5:G5")
                                };
                                //
                                MergeCell mergeCell9 = new MergeCell()
                                {
                                    Reference =
                                   new StringValue("B7:G7")
                                };

                                mergeCells.Append(mergeCell1);
                                mergeCells.Append(mergeCell2);
                                mergeCells.Append(mergeCell3);
                                mergeCells.Append(mergeCell4);
                                mergeCells.Append(mergeCell5);
                                mergeCells.Append(mergeCell6);
                                mergeCells.Append(mergeCell7);
                                mergeCells.Append(mergeCell8);
                                mergeCells.Append(mergeCell9);

                                ///////////////// TABLE BODY FROM 11 ROW INDEX /////////////////////////
                                uint rowIndex = 11;
                                int workId = 1;
                                foreach (Division div in division.divisions)
                                {
                                    Row rowDiv = new Row();
                                    rowDiv.Append(
                                       ConstructCell("", CellValues.String, 2),                                                                                                                   
                                       ConstructCell($"Раздел: {workId.ToString()} , {div.title}" , CellValues.String, 2), // 19/03/2019
                                       ConstructCell("", CellValues.String, 2),
                                       ConstructCell("", CellValues.String, 2),
                                       ConstructCell("", CellValues.String, 2),
                                       ConstructCell("", CellValues.String, 2),
                                       ConstructCell("", CellValues.String, 2));

                                    sheetData.AppendChild(rowDiv);

                                    MergeCell mergeCellDiv = new MergeCell()
                                    {
                                        Reference =
                                   new StringValue("B" + rowIndex.ToString() + ":" + "G" + rowIndex.ToString())
                                    };

                                    mergeCells.Append(mergeCellDiv);

                                    rowIndex++;
                                   

                                    foreach (Work work in div.works)
                                    {
                                        Row rowWork = new Row();
                                        rowWork.Append(
                                           ConstructCell("", CellValues.String, 3),
                                           ConstructCell(workId.ToString(), CellValues.Number, 3),
                                           ConstructCell(work.title, CellValues.String, 3),
                                           ConstructCell(work.units, CellValues.String, 3),
                                           ConstructCell(work.count, CellValues.String, 3),
                                           ConstructCell(work.mtp, CellValues.String, 3),
                                           ConstructCell(work.comment, CellValues.String, 3));

                                           sheetData.AppendChild(rowWork);

                                           rowIndex++;
                                           workId++;
                                    }

                                }


                                Row rowEnd1 = new Row();
                                rowEnd1.Append(
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("Условия стеснённости:", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0));

                                sheetData.AppendChild(rowEnd1);
                                rowIndex++;

                                if (isLookUpNotNull)
                                {
                                    Row rowEnd2 = new Row();
                                    rowEnd2.Height = 80.0;
                                    rowEnd2.Append(
                                       ConstructCell("", CellValues.String, 0),
                                       ConstructCell(lookUpValue, CellValues.String, 7),
                                       ConstructCell("", CellValues.String, 0),
                                       ConstructCell("", CellValues.String, 0),
                                       ConstructCell("", CellValues.String, 0),
                                       ConstructCell("", CellValues.String, 0),
                                       ConstructCell("", CellValues.String, 0));

                                    sheetData.AppendChild(rowEnd2);

                                    MergeCell mergeCellLookUp = new MergeCell()
                                    {
                                        Reference =
                                     new StringValue("B" + rowIndex.ToString() + ":" + "G" + (rowIndex + 1).ToString())
                                    };

                                    mergeCells.Append(mergeCellLookUp);
                                }

                                Row rowEnd3 = new Row();
                                rowEnd3.Append(
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0));

                                sheetData.AppendChild(rowEnd3);


                                Row rowEnd4 = new Row();
                                rowEnd4.Append(
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0));

                                sheetData.AppendChild(rowEnd4);


                                Row rowEnd5 = new Row();
                                rowEnd5.Append(
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell(position3, CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0),
                                   ConstructCell(name3, CellValues.String, 0),
                                   ConstructCell("", CellValues.String, 0));

                                sheetData.AppendChild(rowEnd5);

                                spreadDocument.WorkbookPart.Workbook.Save();
                                spreadDocument.Close();

                                defTitle = defTitle.Replace(".", "");
                                defTitle = defTitle.Replace("-", "");
                                defTitle = defTitle.Replace(":", "");
                                defTitle = defTitle.Replace("_", "");
                                defTitle = defTitle.Replace(" ", "");

                                string fileName = defTitle; 
                                Encoding encoding = Encoding.UTF8;
                                HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                                HttpContext.Current.Response.AddHeader("Content-disposition", "attachment; filename*=UTF-8''" + HttpUtility.UrlEncode(fileName, encoding) + ".xlsx");
                                memoryStream.Position = 0;
                                byte[] arr = memoryStream.ToArray();
                                HttpContext.Current.Response.BinaryWrite(arr);
                                HttpContext.Current.Response.Flush();
                                HttpContext.Current.Response.End();

                            }

                            memoryStream.Close();

                        }

                    }

                });

            }
            catch(Exception ex)
            {
                model.Message = ex.Message;
                Utilities.Log("Дефектная ведомость ", ex.ToString());
            }

        }        
        
    }

}
