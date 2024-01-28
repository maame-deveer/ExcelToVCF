using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using QRCoder;

namespace ExtractAndConvert.GenQrCode
{
    public class VCFToQRCode
    {
        //get image file path
        string imagePath = "";

        //convert .vcf file to qrcode
        public void ConvertToQRCode(string vcfFilePath, string path)
        {
            // Parse the VCF data
            string cardDetails = File.ReadAllText(vcfFilePath);
            string fileName = Path.GetFileNameWithoutExtension(vcfFilePath);
            string qrcode = Path.Combine(path, fileName + ".png");

            //get image path
            imagePath = qrcode;

            // Generate the QR code
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(cardDetails.ToString(), QRCodeGenerator.ECCLevel.M);
            QRCode qrCode = new QRCode(qrCodeData);

            // Convert the QR code to a bitmap image
            Bitmap qrCodeImage = qrCode.GetGraphic(20);

            // Save the QR code image to a file (as PNG)
            using (FileStream stream = new FileStream(qrcode, FileMode.Create))
            {
                qrCodeImage.Save(stream, ImageFormat.Png);
            }
        }

        //get the qrcode image path
        public string QrCodePath()
        {
            return imagePath;
        }
    }
}