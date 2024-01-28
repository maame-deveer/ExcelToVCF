namespace ExtractAndConvert.Parse
{
    //type of data stored in each row depending on the column
    public class WorkSheetRows
    {
        public string? FirstName { get; set; }
        public string? LastName { get; set; }
        public string? Institution { get; set; }
        public List<_Address>? Address {get; set;}
        public string? Position { get; set; }
        public List<PhoneNumber>? Numbers { get; set; }
        public List<_Email>? Emails { get; set; }
        public string? Link { get; set; }

        //to check if row details are being parsed
        /* public override string ToString()
        {
            // Customize the output format here based on your preference
            return $"FirstName: {FirstName}, LastName: {LastName}, Institution: {Institution}, Position: {Position}, Numbers: {FormatList(Numbers)}, Emails: {FormatList(Emails)}, Link: {Link}";
        } */

        // Helper method to format List<T> items as a comma-separated string
        /*private string FormatList<T>(List<T> list)
        {
            if (list == null || list.Count == 0)
            {
                return "N/A";
            }

            return string.Join(", ", list);
        }*/
    }

    public class _Address
    {
        public string? Address { get; set; }
    }

    public class PhoneNumber
    {
        public string? Number { get; set; }
    }

    public class _Email
    {
        public string? EmailAddress { get; set; }
    }
}