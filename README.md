Generate-ppt-using-ASP.Net
==========================

This code will help users to create a ppt of analytic reports which can be used by business or higher management.Saves time in creating presentations manually.


In  .aspx page
<script type="text/javascript">
        function ExportMyChart() {
            var count = 0;
            initiateExport = true;

            document.getElementById("MainContent_hfcount").value = 0;
            document.getElementById("MainContent_hfexportedcount").value = 0;
            for (var chartRef in FusionCharts.items) {

                if (FusionCharts.items[chartRef].exportChart) {
                    count++;
                    FusionCharts.items[chartRef].exportChart();
                }
            }

            document.getElementById("MainContent_hfcount").value = count;

            return false;

        }
       
    </script>
    <script type="text/javascript">
        function FC_Exported(statusObj) {

            var exportedCount = 0;
            if (statusObj.statusCode == 1) {
                exportedCount = document.getElementById("MainContent_hfexportedcount").value;
                exportedCount++;
                document.getElementById("MainContent_hfexportedcount").value = exportedCount;
            }

            if (document.getElementById("MainContent_hfcount").value == document.getElementById("MainContent_hfexportedcount").value) {

                __doPostBack('ctl00$MainContent$lnkExportToPPT', 'lnkExportToPPT_Click');
            }

        }
    </script>

Add two hidden fields:
<asp:HiddenField ID="hfcount" Value="0" runat="server" />
    <asp:HiddenField ID="hfexportedcount" Value="0" runat="server" />

 On Button Call JS function like this:

<asp:ImageButton ID="lnkExportToPPT" runat="server" OnClientClick="javascript:return ExportMyChart();"
                    OnClick="lnkExportToPPT_Click" ImageUrl="~/images/icon-ppt.png" CssClass="exPPT">
                </asp:ImageButton>


On Server Side the Code is

  #region Export
        protected void lnkExportToPPT_Click(object sender, EventArgs e)
        {

            Export();

        }
        public event EventHandler contentCallEvent;
        protected void Export()
        {
            try
            {
                string DirectoryDownload = string.Empty;
                if (WebConfigurationManager.AppSettings["DirectoryDownload"].Count() > 0)
                {
                    DirectoryDownload = WebConfigurationManager.AppSettings["DirectoryDownload"].ToString();
                }
                PPoint.Application objApp;
                objApp = new Microsoft.Office.Interop.PowerPoint.Application();
                PPoint.Presentations objPresSet;
                objPresSet = objApp.Presentations;
                PPoint._Presentation objPres;
                PPoint.Slides objSlides;
                PPoint._Slide objSlide;
                PPoint.TextRange objTextRng, objTextRng2, objTextRng3, objTextRng4, objTextRng5, objFilter1, objFilter2, objFilter3;


                String FileName = "";
                String FilePath = "";

                FileName = "DeviceAggregation.pptx";
                FilePath = DirectoryDownload + "\\DeviceAggregation.pptx";

                objPres = objPresSet.Open(FilePath, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
                objSlides = objPres.Slides;
                var directory = new DirectoryInfo(DirectoryDownload);
                var allowedExtensions = new string[] { ".jpg", ".bmp" };

                var imageFiles = from file in directory.EnumerateFiles("*", SearchOption.AllDirectories)
                                 where allowedExtensions.Contains(file.Extension.ToLower())
                                 select file;


                int Left = 0, Left1 = 0, Left2 = 0;
                foreach (var file in imageFiles)
                {
                    objSlide = objSlides._Index(1);//Add(1, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutObject);
                    objFilter1 = objSlide.Shapes[40].TextFrame.TextRange;
                    objFilter1.Text = ddlCountry.SelectedItem.Text;

                    //objFilter2 = objSlide.Shapes[5].TextFrame.TextRange;
                    //objFilter2.Text = (ViewState["FromUrl"].ToString() == "2" ? "Computers (laptop, convertible, desktop, AIO)" : (ViewState["FromUrl"].ToString() == "3" ? "Laptop + Convertible" : (ViewState["FromUrl"].ToString() == "4" ? "Desktop + AIO" : "0")));

                    objTextRng = objSlide.Shapes[33].TextFrame.TextRange;
                    objTextRng.Text = (ViewState["FromUrl"].ToString() == "2" ? "Computers " : (ViewState["FromUrl"].ToString() == "3" ? "Laptop + Convertible" : (ViewState["FromUrl"].ToString() == "4" ? "Desktop + AIO" : "0")));
                    objTextRng.Font.Size = 12;
                    objTextRng2 = objSlide.Shapes[34].TextFrame.TextRange;
                    objTextRng2.Text = ViewState["OS1"].ToString();
                    objTextRng2.Font.Size = 12;
                    objTextRng3 = objSlide.Shapes[35].TextFrame.TextRange;
                    objTextRng3.Text = ViewState["OS2"].ToString();
                    objTextRng3.Font.Size = 12;
                    objTextRng4 = objSlide.Shapes[36].TextFrame.TextRange;
                    objTextRng4.Text = ViewState["OS3"].ToString();
                    objTextRng4.Font.Size = 12;
                    objTextRng5 = objSlide.Shapes[37].TextFrame.TextRange;
                    objTextRng5.Text = ViewState["OS4"].ToString();
                    objTextRng5.Font.Size = 12;


                    Microsoft.Office.Interop.PowerPoint.Shape shape = objSlide.Shapes[20];
                    if (file.FullName.Contains("Overall Satisfaction"))
                    {
                        if (file.FullName.Contains(ViewState["OS1"].ToString()))
                        {
                            shape = objSlide.Shapes[19];
                        }
                        else if (file.FullName.Contains(ViewState["OS2"].ToString()))
                        {
                            shape = objSlide.Shapes[20];
                        }
                        else if (file.FullName.Contains(ViewState["OS3"].ToString()))
                        {
                            shape = objSlide.Shapes[21];
                        }
                        else if (file.FullName.Contains(ViewState["OS4"].ToString()))
                        {
                            shape = objSlide.Shapes[22];
                        }
                        else
                        {
                            shape = objSlide.Shapes[18];
                        }
                        //shape.Top = 130;
                        //Left = Left + 120;
                        //shape.Left = Left;
                    }

                    else if (file.FullName.Contains("Device OS satisfaction"))
                    {

                        if (file.FullName.Contains(ViewState["OS1"].ToString()))
                        {
                            shape = objSlide.Shapes[24];
                        }
                        else if (file.FullName.Contains(ViewState["OS2"].ToString()))
                        {
                            shape = objSlide.Shapes[25];
                        }
                        else if (file.FullName.Contains(ViewState["OS3"].ToString()))
                        {
                            shape = objSlide.Shapes[26];
                        }
                        else if (file.FullName.Contains(ViewState["OS4"].ToString()))
                        {
                            shape = objSlide.Shapes[27];
                        }
                        else { shape = objSlide.Shapes[23]; }
                        //shape.Top = 240;
                        //Left1 = Left1 + 120;
                        //shape.Left = Left1;
                    }
                    else if (file.FullName.Contains("Touch"))
                    {
                        if (file.FullName.Contains(ViewState["OS1"].ToString()))
                        {
                            shape = objSlide.Shapes[29];
                        }
                        else if (file.FullName.Contains(ViewState["OS2"].ToString()))
                        {
                            shape = objSlide.Shapes[30];
                        }
                        else if (file.FullName.Contains(ViewState["OS3"].ToString()))
                        {
                            shape = objSlide.Shapes[31];
                        }
                        else if (file.FullName.Contains(ViewState["OS4"].ToString()))
                        {
                            shape = objSlide.Shapes[32];
                        }
                        else
                        {
                            shape = objSlide.Shapes[28];
                        }
                        //shape.Top = 380;
                        //Left2 = Left2 + 120;
                        //shape.Left = Left2;
                    }
                    //shape.Width = 90;
                    //shape.Height = 90;
                    objSlide.Shapes.AddPicture(file.FullName, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, shape.Left, shape.Top, shape.Width, shape.Height);
                    NAR(shape);

                    NAR(objTextRng);
                    NAR(objSlide);
                }

                objPres.SaveAs(FilePath.Replace(".pptx", "_New"), PPoint.PpSaveAsFileType.ppSaveAsPresentation, MsoTriState.msoTrue);
                NAR(objSlides);
                objPres.Close();
                NAR(objPres);
                NAR(objPresSet);
                objApp.Quit();
                NAR(objApp);

                foreach (var file in imageFiles)
                {
                    System.IO.File.Delete(file.FullName);
                }

                System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
                response.ClearContent();
                response.Clear();
                response.ContentType = "application/x-mspowerpoint";
                response.AddHeader("Content-Disposition", "attachment; filename=" + FileName.Replace(".pptx", "_New.ppt") + ";");
                response.TransmitFile(FilePath.Replace(".pptx", "_New.ppt"));
                response.Flush();
                response.End();
                objPres.Close();
                objApp.Quit();

            }
            catch (Exception ex)
            {
                GlobalObject.WriteError(ex.Message);
                GlobalObject.ShowMsg(ex.Message);
            }
            finally
            {
            }

        }
        private void NAR(object o)
        {
            try
            {
                Marshal.FinalReleaseComObject(o);
            }
            catch
            {
            }
            finally
            {
                o = null;
            }
        }
        #endregion



Points To Remember:
1 In C Drive make a folder with “Download” name and it must have all the rights including Admin rights.
2. Open Visual studio with as administrator.
3. I am Sending you a project Fc_Export. Host it on IIS. And write the path in  Web Cofig llke this

<appSettings>
    
    <add key="DirectoryDownload" value="C:\\Download"/>
    <add key="ImagesPath" value="http://localhost/Export_.Net/Export_.NET/ExportHandler/FCExporter.aspx"/>
  </appSettings>

Also I am sending you the swf file which should be placed in your project.


The Proejct which I am sen ding you must be in following path

C:\inetpub\wwwroot\Export_.Net
