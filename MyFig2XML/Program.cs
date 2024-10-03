

using System.Net;
using System.Net.Security;
using HtmlAgilityPack;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using SeleniumUndetectedChromeDriver;
using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using System;
using MyFig2XML;

ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, errors) =>
{
    return true;
};



//  Check datafiles exist
string programPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
programPath = programPath.Substring(0, programPath.Length - 14);

if (!File.Exists(programPath + "/last.txt") | !File.Exists(programPath + "/tempicon.png"))
{
    File.Create(programPath + "/last.txt");
    File.Create(programPath + "/tempicon.png");
    Console.WriteLine("Please restart to continue...");
    Console.ReadLine();
    return;
}

var baseDir = AppDomain.CurrentDomain.BaseDirectory;

float conversionRate;

//ChromeDriver webDriver = new ChromeDriver(chromeDriverDirectory: baseDir);

ChromeOptions options = new();

options.AddArgument("minimize_me_:3");

//options.PageLoadTimeout = TimeSpan.FromSeconds(3);

UndetectedChromeDriver webDriver = UndetectedChromeDriver.Create(options, null, driverExecutablePath : await new ChromeDriverInstaller().Auto());//  baseDir + "/chromedriver.exe");

using (IWebDriver driver = webDriver)
{
    //  start program
    Init();
}

Console.WriteLine("\n\n###################################################");
Console.WriteLine("All done!");
Console.WriteLine("Prices exported to: " + programPath + "\\Data.xlsl");
Console.WriteLine("###################################################\n\n");
Console.ReadLine();

void Init()
{
    Console.WriteLine("____________________________________");
    Console.WriteLine("Enter MyFigureCollection username:");
    Console.WriteLine("(Leave empty to re-use last entry)");
    
    string link1 = "https://myfigurecollection.net/?mode=view&username=";
    string link2 = "&tab=collection&status=";
    string link3 = "&current=keywords&rootId=-1&categoryId=-1&output=2&sort=";

    string category = "", ascDesc = "", grouping = "";

    string name = Console.ReadLine();
    string url = link1 + name + link2;
    if (name != string.Empty)
    {
        Console.WriteLine("Enter collection category:");
        Console.WriteLine("wished = 0, ordered = 1, owned = 2, favorites = 3");
        int page = 0;

        try
        {
            page = Convert.ToInt32(Console.ReadLine());
            page = page < 3 ? page : 3;
            Console.WriteLine("Selected: " + page);
        }
        catch { Console.WriteLine("Defaulting to: 0 (wished)"); }

        url += page + link3;

        //category&order=asc&_tb=user&page=

        Console.WriteLine("Sort by which method?");
        Console.WriteLine("category = 0, date = 1, popularity = 2, price = 3, wishability = 4, activity = 5");
        int cat = 0;

        try
        {
            cat = Convert.ToInt32(Console.ReadLine());
            cat = cat < 5 ? cat : 5;
        }
        catch { Console.WriteLine("Defaulting to: 0 (category)"); }

        switch (cat)
        {
            case 0: { category = "category"; break; }
            case 1: { category = "date"; break; }
            case 2: { category = "popularity"; break; }
            case 3: { category = "price"; break; }
            case 4: { category = "wishability"; break; }
            case 5: { category = "activity"; break; }
        }

        url += (category);

        Console.WriteLine("Ascending or descending?");
        Console.WriteLine("ascending = 0, descending = 1");

        int updown = 0;

        try
        {
            updown = Convert.ToInt32(Console.ReadLine());
            updown = updown < 1 ? updown : 1;
        }
        catch { Console.WriteLine("Defaulting to: 0 (ascending)"); }

        ascDesc = updown == 1 ? "desc" : "asc";

        Console.WriteLine("Group by which method?");
        Console.WriteLine("none = 0, origins = 1, characters = 2, companies = 3, artists = 4, classifications = 5, scales = 6, releaseDates = 7, wishabilities = 8");
        
        int group = 0;

        try
        {
            group = Convert.ToInt32(Console.ReadLine());
            group = group < 8 ? group : 8;
        }
        catch { Console.WriteLine("Defaulting to: 0 (origins)"); }

        switch (group)
        {
            case 0: { grouping = ""; break; }
            case 1: { grouping = "origins"; break; }
            case 2: { grouping = "characters"; break; }
            case 3: { grouping = "companies"; break; }
            case 4: { grouping = "artists"; break; }
            case 5: { grouping = "classifications"; break; }
            case 6: { grouping = "scales"; break; }
            case 7: { grouping = "releaseDates"; break; }
            case 8: { grouping = "wishabilities"; break; }
        }

        url += (grouping == "" ? "" : ("&groupBy=" + grouping)) + "&order=" + ascDesc + "&_tb=user&page=";
    }

    //  use last url
    if (name == string.Empty)
    {

        if (File.Exists(programPath + "/last.txt"))
            url = File.ReadAllText(programPath + "/last.txt");
        else
            return;
    }

    File.WriteAllText(programPath + "/last.txt", url);

    //  get conversion rate
    var rates = FetchWebData("https://v6.exchangerate-api.com/v6/f12b66341ce2c43b744559c8/latest/USD");

    conversionRate = float.Parse(rates.Split("\"JPY\":")[1].Split(",")[0]);

    Console.WriteLine("Current JPY -> USD conversion rate is: " + conversionRate);

    //  begin
    if (url.ToLower().Contains("myfigurecollection.net"))
        MainProcess(url);
}

void MainProcess(string url)
{
    using (var workbook = new XLWorkbook())
    {
        //  setup OpenXML
        var worksheet = workbook.Worksheets.Add("Figure Data");

        int row = 1;

        worksheet.Cell(row, 1).Value = "Name";
        worksheet.Cell(row, 2).Value = "Nin-Nin-Game";
        worksheet.Cell(row, 3).Value = "Solaris Japan";
        worksheet.Cell(row, 4).Value = "USD->JPY: " + conversionRate;

        worksheet.Column(1).Width = 130;
        worksheet.Column(2).Width = 15;
        worksheet.Column(3).Width = 15;

        workbook.SaveAs("Data.xlsx");

        var alwaysContinue = false;
        var end = false;

        for(int p = 1; p < 9999; p++)
        {
            if(p > 1 & !alwaysContinue)
            {
                Console.WriteLine("Continue to the next page, or stop here?");
                Console.WriteLine("y - continue     n - stop     a - always continue");
                var ask = Console.ReadLine();
                switch (ask)
                {
                    case "n": { end = true; alwaysContinue = true; break; }
                    case "a": { alwaysContinue = true; break; }
                }

                //if(end) break;
            }

            //  Get collection webpage
            string data = FetchWebData(url + p);

            //Console.WriteLine($"{data}");

            var itemSplit = data.Split("class=\"item-icon\"");
            
            if ((itemSplit.Length - 1) <= 0)
                break;

            Console.WriteLine("----------------------------");
            Console.WriteLine("P A G E: "+ p +"    L E N G T H: " + (itemSplit.Length - 1));
            Console.WriteLine("----------------------------");

            for (int i = 0; i < itemSplit.Length; i++)
            {
                if (i == 0) continue;
                //  1   =   item link
                //  5   =   small image
                //  7   =   item name

                row++;

                var subData = itemSplit[i].Split("\"");

                if (subData.Length < 7) continue;

                string link = subData[1], icon = subData[5], name = subData[7], solaris = "---", ninin = "---", sLink = "", nLink = "";

                var itemPage = string.Empty;
                if(!end)
                    itemPage = FetchWebData("https://myfigurecollection.net" + link);

                var sellers = itemPage.Split("icon icon-diamond");

                if (sellers.Length > 1)
                {
                    var partnerPrefix = "https://myfigurecollection.net/?_tb=partner&amp;mode=goto&amp;partnerId=";
                    var filteredLinks = sellers[1].Split("result-actions");

                    List<string> sortedLinks = new();

                    foreach (string fLink in filteredLinks)
                    {
                        List<string> sellerLinks = new();

                        if (fLink.Length == 1) continue;

                        sellerLinks.Add(fLink.Split(partnerPrefix)[1]);

                        for (int l = 0; l < sellerLinks.Count; l++)
                        {
                            if (i % 2 != 0) continue;

                            var id = -1;

                            if (int.TryParse(sellerLinks[l].Substring(0, 2), out id))
                            { }
                            else { continue; }

                            //  49 = solaris    52 = nin-nin
                            if (id != 49 & id != 52) continue;

                            var finalLink = (partnerPrefix + sellerLinks[l].Split("\"")[0]).Replace(';', '&');

                            foreach (string testLink in sortedLinks)
                            {
                                if (testLink == finalLink)
                                {
                                    //Console.WriteLine("Skipping link, identical");
                                    goto end_of_loop;
                                }
                            }

                            sortedLinks.Add(finalLink);

                            var partnerData = FetchWebData(finalLink);

                            try
                            {
                                if (id == 49)
                                {
                                    solaris = SolarisPrice(partnerData);
                                    sLink = finalLink;
                                }
                                else if (id == 52)
                                {
                                    ninin = NinNinPrice(partnerData);
                                    nLink = finalLink;
                                }
                            }
                            catch { }

                        end_of_loop:;
                        }
                    }
                }

                Console.WriteLine("Name: " + name);
                //Console.WriteLine("Link: " + link);
                Console.WriteLine("Icon: " + icon);
                //DisplayIcon(icon);
                Console.WriteLine("Nin-Nin price: " + ninin);
                Console.WriteLine("Solaris price: " + solaris);
                Console.WriteLine("( " + i + " / " + (itemSplit.Length - 1) + " ) page: " + p);

                worksheet.Cell(row, 1).Value = name;
                worksheet.Cell(row, 1).SetHyperlink(new(@"https://myfigurecollection.net" + link));
                worksheet.Cell(row, 2).Value = ninin;
                if(ninin != "---")
                    worksheet.Cell(row, 2).SetHyperlink(new(@nLink));
                worksheet.Cell(row, 3).Value = solaris;
                if(solaris != "---")
                    worksheet.Cell(row, 3).SetHyperlink(new(@sLink));

                try {
                    workbook.Save();
                }
                catch {
                    Console.WriteLine("!   !   !   !   !   !   !   !   !   !   !   !   !   !   !");
                    Console.WriteLine("Couldn't save the document!!! :c");
                    Console.WriteLine("Make sure you don't have it open");
                    Console.WriteLine("!   !   !   !   !   !   !   !   !   !   !   !   !   !   !");
                    Console.ReadLine();
                }
                Console.WriteLine("----------------------------");
            }
        }
    }
}

string SolarisPrice(string data)
{
    var isolatedData = data.Split("h5--body product__price")[1];

    isolatedData = isolatedData.Split("currency&quot;>")[1];

    var usd = isolatedData.Substring(0, 1) == "$";

    isolatedData = isolatedData.Split("</span>")[1];

    isolatedData = isolatedData.Split("\"")[0];

    float price = -1;

    float.TryParse(isolatedData, out price);

    if (price != -1 & !usd)
        price /= conversionRate;

    isolatedData = price == -1 ? "---" : price.ToString();

    return isolatedData;
}

string NinNinPrice(string data)
{
    var isolatedData = data.Split("id=\"our_price_display\"> ")[1];

    var usd = isolatedData.Substring(0, 3) == "US$";

    isolatedData = isolatedData.Substring(3, isolatedData.Length - 3);
    isolatedData = isolatedData.Split("<")[0];

    float price = -1;

    float.TryParse(isolatedData, out price);

    if(price != -1 & !usd)
        price /= conversionRate;

    isolatedData = price == -1 ? "---" : price.ToString();

    return isolatedData;
}

string FetchWebData(string link)
{
    Console.WriteLine("Fetching: " + link);

    try
    {
        // Navigate to the desired webpage
        webDriver.GoToUrl(link);

        // Wait for the page to load completely
        System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3).Milliseconds);
        //System.Threading.Thread.Sleep(300); // Adjust as necessary (was 2000)

        // Get the page source after JavaScript execution
        string pageSource = webDriver.PageSource;

        // Load the page source into Html Agility Pack
        HtmlDocument doc = new HtmlDocument();
        doc.LoadHtml(pageSource);

        // Now you can manipulate or extract data from 'doc'
        return doc.Text;
    }
    catch
    {
        Console.WriteLine("!   !   !   !   !   !   !   !   !   !   !   !   !   !   !");
        Console.WriteLine("There was an error or timeout with this webpage!!! :c");
        Console.WriteLine("You can continue if you want, but there may be missing data or issues");
        Console.WriteLine("!   !   !   !   !   !   !   !   !   !   !   !   !   !   !");
        Console.ReadLine();
        return "";
    }
}

void DisplayIcon(string link)
{
    using (WebClient client = new WebClient())
    {
        var iconPath = programPath + @"\tempicon.png";
        client.DownloadFile(new Uri(link), iconPath);

        ConsoleGraphics.Render(iconPath);
    }
}