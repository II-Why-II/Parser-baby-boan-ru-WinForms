using ImageMagick;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Parser_baby_boan_ru
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Text = "Parser for boan-baby.ru";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public int PictureCounter = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            string url = textBox1.Text;
            string BaseDirectoryPath = textBox2.Text;
            string imageFormat = comboBox1.Text;
            string DirectoryPath = BaseDirectoryPath + url.Replace("https://www.boan-baby.ru/", "\\").Replace("product/", "");

            if (imageFormat == string.Empty || imageFormat == null)
            {
                imageFormat = "jpeg";
                MessageBox.Show("Не указан формат сохраняемых изображений. По умолчанию задан .jpeg");
            }
            
            var urlListWithAllColors = getUrlsAboutAllColorsOrGetNullIfException(url);
            _ = 1;
            foreach (var urlAboutColor in urlListWithAllColors)
            {
                string subDirectoryPath = DirectoryPath + urlAboutColor.Replace("https://www.boan-baby.ru/", "\\").Replace("product/", "");
                _ = 1;
                checkOnExistsTheFolders(subDirectoryPath);

                string htmlThisPage = getPageHtmlOrNullIfExceprion(urlAboutColor);

                ProductInfo productModel = ParseThePageAndReturnData(htmlThisPage);

                productModel.Url = urlAboutColor;

                productModel.DirectoryWithPicturesPath = subDirectoryPath;
                
                productModel.PictureUrls = getPictureUrlsOrNull(htmlThisPage);

                productModel.YoutubeHrefs = getYoutubeUrlsOrNull(urlAboutColor);
                _ = 1;
                int numberOfImage = 0;
                foreach (var pictureUrl in productModel.PictureUrls)
                {
                    var picturePath = getDictionaryWithDownloadPucturesToComputerOrNull(pictureUrl, subDirectoryPath, Convert.ToString(numberOfImage), imageFormat);
                    numberOfImage++;

                    foreach(var path in picturePath)
                    {
                        if (path.Key == "Origin picture")
                            productModel.PathsOriginalImages.Add(path.Value);

                        if (path.Key == "File for edit")
                            productModel.PathsImagesForEdit.Add(path.Value);
                    }
                }

                PictureCounter += productModel.PathsImagesForEdit.Count;
                label5.Text = PictureCounter.ToString();

                foreach (var picturePath in productModel.PathsImagesForEdit) 
                {
                    resizeImageAndSaveIt(picturePath, subDirectoryPath + "\\Edited images");
                }
                
                saveDataToExcel(BaseDirectoryPath, productModel);
                _ = 1;
            }
            _ = 1;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string BaseDirectoryPath = textBox2.Text + "\\Downloaded images";
            string url = textBox3.Text;
            string fileOriginalFormat = url.Split('.').Last();
            string imageFormat = comboBox1.Text;
            string imageName = url.Replace("https://", "").Replace("www.", "").Replace("boan-baby.ru/", "").Replace("product/", "").Replace("." + fileOriginalFormat, "").Replace("/", "-");

            if (imageFormat == string.Empty || imageFormat == null)
            {
                imageFormat = "jpeg";
                MessageBox.Show("Не указан формат сохраняемых изображений. По умолчанию задан .jpeg");
            }

            checkOnExistsTheFolders(BaseDirectoryPath);

            try
            {
                var dictionaryImgPaths = getDictionaryWithDownloadPucturesToComputerOrNull(url, BaseDirectoryPath, imageName, imageFormat);
                List<string> originFile = new List<string>();
                List<string> filesForEdit = new List<string>();
                foreach (var path in dictionaryImgPaths)
                {
                    if (path.Key == "Origin picture")
                        originFile.Add(path.Value);

                    if (path.Key == "File for edit")
                        filesForEdit.Add(path.Value);
                    PictureCounter++;
                    label5.Text = PictureCounter.ToString();
                }

                foreach (var picPaths in filesForEdit)
                    resizeImageAndSaveIt(picPaths, BaseDirectoryPath + "\\Edited images");
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error with picture in origin method: " + ex.Message);
            }
        }

        private void checkOnExistsTheFolders(string subDirectoryPath) 
        {
            string originImagePath = subDirectoryPath + "\\Original images";
            string forEditPaths = subDirectoryPath + "\\Images for edit";
            string editedFilePaths = subDirectoryPath + "\\Edited images";
            string subEditedNoReducedFiles = editedFilePaths + "\\No reduced images";
            string subEditedReducedFiles = editedFilePaths + "\\Reduced images";


            if (!Directory.Exists(subDirectoryPath))
                Directory.CreateDirectory(subDirectoryPath);

            if (!Directory.Exists(originImagePath))
                Directory.CreateDirectory(originImagePath);

            if (!Directory.Exists(forEditPaths))
                Directory.CreateDirectory(forEditPaths);

            if (!Directory.Exists(editedFilePaths))
                Directory.CreateDirectory(editedFilePaths);

            if (!Directory.Exists(subEditedNoReducedFiles))
                Directory.CreateDirectory(subEditedNoReducedFiles);

            if (!Directory.Exists(subEditedReducedFiles))
                Directory.CreateDirectory(subEditedReducedFiles);
        }
        private string getPageHtmlOrNullIfExceprion(string url)
        {
            try
            {
                string pageInString = null;
                WebRequest request = WebRequest.Create(url);
                WebResponse response = request.GetResponse();

                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream))
                    pageInString = reader.ReadToEnd();
                return pageInString;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loading error: " + ex.Message);
                return null;
            }
        }
        private string getPathSavingFile(string path, string fileName, string fileFormat)
        {
            return path + $"\\{fileName}.{fileFormat}";
        }

        private ProductInfo ParseThePageAndReturnData(string html)
        {
            ProductInfo productInfo = new ProductInfo();
            try
            {
                HtmlAgilityPack.HtmlDocument documentAboutDiv = new HtmlAgilityPack.HtmlDocument();
                documentAboutDiv.LoadHtml(html);
                try
                {
                    productInfo.Name = documentAboutDiv.DocumentNode.SelectSingleNode("//div[@class='product_page_name_block']/h1").InnerText.Replace("&bull;", "").Replace("&nbsp;", "").Replace("&mdash;", "");
                }
                catch (Exception ex) { }

                try
                {
                    productInfo.Color = documentAboutDiv.DocumentNode.SelectSingleNode("//div[@class='product_variant_item active']/div[@class='product_variant_name']").InnerText.Replace("\t", "").Replace("\n", "").Replace("&bull;", "").Replace("&nbsp;", "").Replace("&mdash;", "");
                }
                catch (Exception ex) { }
                try
                {
                    productInfo.ShortDescription = documentAboutDiv.DocumentNode.SelectSingleNode("//div[@id='product_short_desc_wrap']/div[@class='product_short_desc']").InnerText.Replace("&bull;", "").Replace("&nbsp;", "").Replace("&mdash;", "");
                }
                catch (Exception ex) { }
                try
                {
                    string descriptionHtml = documentAboutDiv.DocumentNode.SelectSingleNode("//div[@itemprop='description']").InnerHtml;

                    productInfo.FullDescription = documentAboutDiv.DocumentNode.SelectSingleNode("//div[@itemprop='description']").InnerText.Replace("&bull;", "").Replace("&nbsp;", "").Replace("&mdash;", "");

                    var descriptionHtmlSeparatedByH3 = descriptionHtml.Split(new string[] { "<h3>" }, StringSplitOptions.RemoveEmptyEntries);

                    _ = 1;
                    foreach (var h3 in descriptionHtmlSeparatedByH3)
                    {
                        string h3string = h3;
                        if (h3 != descriptionHtmlSeparatedByH3[0])
                            h3string = "<h3>" + h3;

                        HtmlAgilityPack.HtmlDocument h3Document = new HtmlAgilityPack.HtmlDocument();
                        h3Document.LoadHtml(h3string);

                        if (h3 == descriptionHtmlSeparatedByH3[0])
                            productInfo.LongDescription = getPResultListIntoH3(h3Document);

                        if (h3string.Contains("<strong>Характеристики</strong>"))
                            productInfo.Characteristics = getPResultListIntoH3(h3Document);

                        if (h3string.Contains("<strong>Габариты</strong>"))
                            productInfo.SizeList = getPResultListIntoH3(h3Document);

                        _ = 1;
                    }
                }
                catch (Exception ex) { }
                _ = 1;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Parsing error: " + ex.Message);
            }
            return productInfo;
        }
        private string getPResultListIntoH3(HtmlAgilityPack.HtmlDocument h3)
        {
            string resultString = string.Empty;

            var longDescriptionNodes = h3.DocumentNode.SelectNodes("//p");

            foreach (var node in longDescriptionNodes)
            {
                var lineSeparation = node.InnerText.Replace("&bull;", "").Replace("&nbsp;", "").Replace("&mdash;", "").Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lineSeparation)
                    resultString += line + "\n";
            }
            _ = 1;
            return resultString;
        }

        private List<string> getUrlsAboutAllColorsOrGetNullIfException(string url)
        {
            var pageHtmlInString = getPageHtmlOrNullIfExceprion(url);
            if (pageHtmlInString != null)
            {
                try
                {
                    HtmlAgilityPack.HtmlDocument docAboutBasePage = new HtmlAgilityPack.HtmlDocument();
                    docAboutBasePage.LoadHtml(pageHtmlInString);

                    List<string> listOfHrefs = new List<string>();
                    try
                    {
                        var activeProductHref = url;
                        listOfHrefs.Add(activeProductHref);

                        var productColors = docAboutBasePage.DocumentNode.SelectNodes("//div[@id='product_variant_block']/a");
                        foreach (var product in productColors)
                        {
                            listOfHrefs.Add("https://www.boan-baby.ru/" + product.Attributes["href"].Value);
                        }
                        _ = 1;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("error loading color pages: " + ex.Message);
                    }
                    return listOfHrefs;
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Error loading page: " + ex.Message);
                    return null;
                }
            }
            else
                return null;
        }
        private List<string> getPictureUrlsOrNull(string html)
        {
            try
            {
                HtmlAgilityPack.HtmlDocument docAboutThisPage = new HtmlAgilityPack.HtmlDocument();
                docAboutThisPage.LoadHtml(html);

                var picHrefs = docAboutThisPage.DocumentNode.SelectNodes("//div[@class='foto_page_wrap']/div/a");

                List<string> pictureUrls = new List<string>();

                foreach (var href in picHrefs)
                {
                    if (!href.Attributes["href"].Value.Contains("youtube"))
                        pictureUrls.Add("https://www.boan-baby.ru/" + href.Attributes["href"].Value);
                }
                pictureUrls.Remove(pictureUrls.Last());
                return pictureUrls;
            }
            catch (Exception ex)
            {
                MessageBox.Show("error get picture urls: " + ex.Message);
                return null;
            }
        }
        private List<string> getYoutubeUrlsOrNull(string html)
        {
            try
            {
                HtmlAgilityPack.HtmlDocument docAboutThisPage = new HtmlAgilityPack.HtmlDocument();
                docAboutThisPage.LoadHtml(html);

                var picHrefs = docAboutThisPage.DocumentNode.SelectNodes("//div[@class='foto_page_wrap']/div/a");

                List<string> youtubeUrls = new List<string>();

                foreach (var href in picHrefs)
                {
                    if (href.Attributes["href"].Value.Contains("youtube"))
                        youtubeUrls.Add("https://www.boan-baby.ru/" + href.Attributes["href"].Value);
                }
                return youtubeUrls;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        
        private Dictionary<string, string> getDictionaryWithDownloadPucturesToComputerOrNull(string url, string directoryPath, string fileName, string fileFormat)
        {
            Dictionary<string, string> dictionaryDownloadedFilesPaths = new Dictionary<string, string>();

            string originImagePath = directoryPath + "\\Original images";
            string editImagePaths = directoryPath + "\\Images for edit";

            //automatic file format copy
            string originFileFormat = url.Split('.').Last();

            string originPicturePath = getPathSavingFile(originImagePath, fileName, originFileFormat);
            string editPicturePath = getPathSavingFile(editImagePaths, fileName, fileFormat);

            try
            {
                _ = 1;
                using (WebClient client = new WebClient())
                {
                    try
                    {
                        client.DownloadFile(url, originPicturePath);
                        dictionaryDownloadedFilesPaths.Add("Origin picture", originPicturePath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error downloading original file:" + ex.Message);
                    }
                    try
                    {
                        client.DownloadFile(url, editPicturePath);
                        dictionaryDownloadedFilesPaths.Add("File for edit", editPicturePath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error downloading file for edit:" + ex.Message);
                    }
                    _ = 1;
                }
                return dictionaryDownloadedFilesPaths;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error load downloading picture:" + ex.Message);
                return null;
            }
        }

        private void resizeImageAndSaveIt(string imagePath, string savingPath)
        {
            string nameOfImage = imagePath.Split('\\').Last();
            string noReducedImagePath = savingPath + "\\No reduced images\\" + nameOfImage;
            string reducedImagePath = savingPath + "\\Reduced images\\" + nameOfImage;

            using (MagickImage whiteSquid = new MagickImage("C:\\data\\special pictures\\White1000x1000.jpeg"))
            {
                MagickImage image = new MagickImage(imagePath);

                MagickGeometry squidSize = new MagickGeometry(1000, 1000);

                // This will resize the image to a fixed size without maintaining the aspect ratio.
                // Normally an image will be resized to fit inside the specified size.
                squidSize.IgnoreAspectRatio = false;

                image.HasAlpha = true;
                image.ColorAlpha(new MagickColor("white"));

                if (image.Width > image.Height)
                    image.Resize(1000, 0);
                else if (image.Width < image.Height)
                    image.Resize(0, 1000);
                else
                    image.Resize(squidSize);

                whiteSquid.Composite(image, Gravity.Center);

                whiteSquid.Write(noReducedImagePath);

                var optimizer = new ImageOptimizer();
                optimizer.LosslessCompress(noReducedImagePath);

                var savedImgInfo = new FileInfo(noReducedImagePath);
                if (savedImgInfo.Length > 256000) //256000 == 250kb.  307200 == 300kb
                {
                    var reducedImgInfo = new FileInfo(reducedImagePath);
                    bool saved = false;
                    for (int quality = 75; saved != true && quality > 0; quality--)
                    {
                        image.Quality = quality;

                        image.Write(reducedImagePath);
                        optimizer.LosslessCompress(reducedImagePath);

                        reducedImgInfo.Refresh();

                        if (reducedImgInfo.Length < 307200)
                            saved = true;
                    }
                    if (saved != true)
                        MessageBox.Show("Error with saving the photo");
                }
                _ = 1;
            }
        }
       
        private void saveDataToExcel(string directoryPath, ProductInfo product)
        {
            FileInfo newFile = new FileInfo(directoryPath + "\\Products.xlsx");

            using (var package = new ExcelPackage(newFile))
            {
                ExcelWorksheet sheet; //= package.Workbook.Worksheets[1]; // 1 in .Net3.5 and .Net 4.0; 0 in .Net core

                if (package.Workbook.Worksheets["Content"] != null)
                    sheet = package.Workbook.Worksheets["Content"];
                else
                    sheet = package.Workbook.Worksheets.Add("Content");
                package.Save();

                int row = 1;
                int column = 1;

                while (sheet.Cells[row, 1].Value != null)
                {
                    row++;
                }

                if (row == 1 && sheet.Cells[1, 1].Value == null)
                {
                    sheet.Cells[row, column++].Value = "Id";
                    sheet.Cells[row, column++].Value = "Url";
                    sheet.Cells[row, column++].Value = "Name";
                    sheet.Cells[row, column++].Value = "Color";
                    sheet.Cells[row, column++].Value = "Short description";
                    sheet.Cells[row, column++].Value = "Fill description";
                    sheet.Cells[row, column++].Value = "LongDescription";
                    sheet.Cells[row, column++].Value = "Characteristics";
                    sheet.Cells[row, column++].Value = "Gabarites";
                    sheet.Cells[row, column++].Value = "Directory With Pictures";
                    column = 1;
                    row++;
                }

                sheet.Cells[row, column++].Value = row - 1;
                sheet.Cells[row, column++].Value = product.Url;
                sheet.Cells[row, column++].Value = product.Name;
                sheet.Cells[row, column++].Value = product.Color;
                sheet.Cells[row, column++].Value = product.ShortDescription;
                sheet.Cells[row, column++].Value = product.FullDescription;
                sheet.Cells[row, column++].Value = product.LongDescription;
                sheet.Cells[row, column++].Value = product.Characteristics;
                sheet.Cells[row, column++].Value = product.SizeList;
                sheet.Cells[row, column++].Value = product.DirectoryWithPicturesPath;

                package.Save();
            }
        }
    }
    class ProductInfo
    {
        public string Url { get; set; }
        public string Name { get; set; }
        public string Color { get; set; }
        public string ShortDescription { get; set; }
        public string LongDescription { get; set; }
        public string FullDescription { get; set; }
        public string Characteristics { get; set; }
        public string SizeList { get; set; }
        public string DirectoryWithPicturesPath { get; set; }

        public List<string> PictureUrls { get; set; } = new List<string>();
        public List<string> YoutubeHrefs { get; set; } = new List<string>();

        public List<string> PathsOriginalImages { get; set; } = new List<string>();
        public List<string> PathsImagesForEdit { get; set; } = new List<string>();
    }
}

