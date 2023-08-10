using ExtractAndConvert.Parse;
using ExtractAndConvert.VCard;
using ExtractAndConvert.GenQrCode;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Text.RegularExpressions;

namespace ExtractAndConvert.SheetToVCF
{
    public class WorkSheetToVCF
    {
        bool completed = false;

        public void ToVCF(string worksheet, string folderPath)
        {
            ParseWorkSheet extract = new ParseWorkSheet();
            VCFToQRCode qrCode = new VCFToQRCode();
            List<WorkSheetRows> extractedData = extract.ExtractDataFromWorkSheet(worksheet);
            bool isFirstRow = true; // Flag to track the first row

            //extract data if it is not null
            if (extractedData != null && extractedData.Any())
            {
                //Access the extracted data
                for (int i = 0; i < extractedData.Count; i++)
                {
                    // Skip parsing the first row
                    if (isFirstRow)
                    {
                        isFirstRow = false;
                        continue;
                    }

                    NewContact contact = new NewContact()
                    {
                        FirstName = extractedData[i].FirstName,
                        LastName = extractedData[i].LastName,
                        Organization = extractedData[i].Institution,
                        Addresses = new List<Address>()
                        {
                            { new Address(){ Location = " "} }
                        },
                        Title = extractedData[i].Position,
                        Phones = new List<Phone>(),
                        Emails = new List<Mail>(),
                        Websites = new List<Website>()
                        {
                            { new Website(){ Link= " "} }
                        }
                    };

                    // Check if extractedData[i].Numbers is not null and contains elements
                    if (extractedData[i].Numbers != null && extractedData[i].Numbers.Any())
                    {
                        foreach (var item in extractedData[i].Numbers)
                        {
                            contact.Phones.Add(new Phone { Number = item?.Number ?? "" });
                        }
                    }

                    // Check if extractedData[i].Emails is not null and contains elements
                    if (extractedData[i].Emails != null && extractedData[i].Emails.Any())
                    {
                        foreach (var item in extractedData[i].Emails)
                        {
                            contact.Emails.Add(new Mail { Value = item?.EmailAddress ?? "" });
                        }
                    }


                    //save 
                    string vcardcontents = CreateVCard.VCard(contact);
                    string fileName = $"{contact.FirstName}_{contact.LastName}({contact.Organization}).vcf";
                    string SaveFilePath = Regex.Replace(Path.Combine(folderPath, fileName), @"\s", "");
                    System.IO.File.WriteAllText(SaveFilePath, vcardcontents);
                }
                completed = true;
            }
        }

        public bool IsCompleted()
        {
            return completed;
        }
    }
}
